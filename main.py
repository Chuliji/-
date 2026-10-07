# -*- coding: utf-8 -*-
"""自动注册本地 _deps 依赖（pywin32 等）"""
import os as _os, sys as _sys


def _get_deps_dir():
    """获取 _deps 目录路径，兼容开发模式和 PyInstaller 打包。"""
    if hasattr(_sys, "_MEIPASS"):
        _p = _os.path.join(_sys._MEIPASS, "_deps")
        if _os.path.isdir(_p):
            return _p
    _exe_dir = _os.path.dirname(_os.path.abspath(_sys.executable))
    _p = _os.path.join(_exe_dir, "_deps")
    if _os.path.isdir(_p):
        return _p
    _p = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), "_deps")
    if _os.path.isdir(_p):
        return _p
    return None


_deps = _get_deps_dir()
if _deps:
    for _sub in ("", "win32", "win32lib"):
        _p = _os.path.join(_deps, _sub) if _sub else _deps
        _lib = _os.path.join(_deps, "win32", "lib")
        for _d in (_p, _lib):
            if _os.path.isdir(_d) and _d not in _sys.path:
                _sys.path.insert(0, _d)
    _dll_dir = _os.path.join(_deps, "pywin32_system32")
    if _os.path.isdir(_dll_dir):
        try:
            if hasattr(_os, "add_dll_directory"):
                _os.add_dll_directory(_dll_dir)
            _os.environ["PATH"] = _dll_dir + _os.pathsep + _os.environ.get("PATH", "")
        # 启动期环境配置只做尽力而为，失败不影响主流程
        # noinspection PyBroadException
        except Exception:
            pass

from typing import Any, Dict, Iterable, List, Match, Optional, Tuple

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.opc.part import Part
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.table import Table
import lxml.etree as etree
import os
import re
import sys
import time
import zipfile
import shutil
import multiprocessing
import copy


def _set_run_font(run, font_name='SimSun', size=Pt(10.5)):
    """设置 run 的中文字体和字号。

    python-docx 的 Font 类不直接暴露 eastAsia 属性，
    需要通过底层 XML 设置 w:eastAsia 属性。
    """
    run.font.name = font_name
    run.font.size = size
    rPr = run._element.get_or_add_rPr()  # type: ignore[attr-defined]
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        from docx.oxml import OxmlElement
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:eastAsia'), font_name)

def _get_template_path() -> str:
    """获取模板文件路径：PyInstaller 打包后从 sys._MEIPASS 取，开发时用当前目录。"""
    if hasattr(sys, "_MEIPASS"):
        p = os.path.join(sys._MEIPASS, "原始记录模版2.0.docx")
        if os.path.exists(p):
            return p
    return "原始记录模版2.0.docx"


TEMPLATE_FILE = _get_template_path()
REPORTS_FOLDER = "检验报告"
OUTPUT_FOLDER = "原始记录汇总"
EXTRACTED_FOLDER = "_已解压"  # 临时解压目录（放在检验报告文件夹下，处理完自动删除）
LOG_FOLDER = os.path.join("终端结果", "main运行日志")   # 每次运行的终端输出都存这里


def _is_valid_report_docx(filename):
    """判断一个 .docx 是否为有效的检验报告。
    只排除临时锁文件(~$)、合格证等非报告文件，其他 .docx 一律放行。
    """
    if not filename.lower().endswith('.docx'):
        return False
    if filename.startswith('~$'):
        return False
    if '合格证' in filename:
        return False
    return True


def _is_report_input(filename):
    """识别任何可作为检验报告输入的文件（.docx/.doc/.wps/.docm）。"""
    if filename.startswith('~$'):
        return False
    if '合格证' in filename:
        return False
    return filename.lower().endswith(('.docx', '.doc', '.wps', '.docm'))


def _list_report_inputs(folder, recursive=True):
    """列出目录下所有报告输入文件，返回相对 folder 的路径（sorted）。"""
    out = []
    if recursive:
        for root, _dirs, files in os.walk(folder):
            for f in files:
                if _is_report_input(f):
                    out.append(os.path.relpath(os.path.join(root, f), folder))
    else:
        out = [f for f in os.listdir(folder) if _is_report_input(f)]
    return sorted(out)


class _Tee:
    """同时写终端和日志文件的 stdout 代理。"""
    def __init__(self, *streams):
        self.streams = streams

    def write(self, data):
        for s in self.streams:
            s.write(data)
            try:
                s.flush()
            except OSError:
                pass

    def flush(self):
        for s in self.streams:
            try:
                s.flush()
            except OSError:
                pass

    def close(self):
        for s in self.streams:
            if s is not sys.stdout:
                try:
                    s.close()
                except OSError:
                    pass


def _build_log_filename(all_filenames):
    """从文件名列表里提取报告编号范围，生成日志文件名。

    例: ['B-WTQC20261001粤SY2760检验报告.docx', ..., 'B-WTQC20261041粤SD6570检验报告.docx']
    -> 'B-WTQC20261001-1041.log'
    """
    prefixes = set()
    suffix_nums = []
    for fn in all_filenames:
        # 匹配 B-WTQC 或 B-ZXQC 开头 + 一串数字
        m = re.search(r'(B-[A-Z]+)(\d+)', fn)
        if m:
            prefixes.add(m.group(1))
            suffix_nums.append(int(m.group(2)))

    if not suffix_nums:
        return f"run_{time.strftime('%Y%m%d_%H%M%S')}.log"

    prefix = sorted(prefixes)[0] if len(prefixes) == 1 else "B-ALL"
    lo, hi = min(suffix_nums), max(suffix_nums)
    lo_str = str(lo)
    hi_str = str(hi)
    # 如果只有一个编号，就只写一个
    if lo == hi:
        return f"{prefix}{lo_str}.log"
    # 前缀相同的部分只保留一次
    common_len = 0
    for a, b in zip(lo_str, hi_str):
        if a == b:
            common_len += 1
        else:
            break
    suffix = hi_str[common_len:]
    return f"{prefix}{lo_str}-{suffix}.log"


def _table_grid(table: Table) -> Tuple[Any, int]:
    """一次性物化整张表格的网格，返回 (cells 扁平列表, 列数)。

    python-docx 的 table.cell(r, c) 每次调用都会重建整张网格并解析所有合并单元格
    (vMerge/gridSpan)，在多层循环里反复调用会极慢。这里先算一次，之后直接按下标取。
    """
    # python-docx 未公开网格 API，直接返回内部成员（仅在物化时触碰一次）
    # noinspection PyProtectedMember
    return table._cells, table._column_count


def _cell_text(grid: Tuple[Any, int], row: int, col: int) -> str:
    """从已物化的网格里按 (row, col) 取单元格文本。"""
    cells, col_count = grid
    try:
        return cells[row * col_count + col].text.strip()
    except (IndexError, AttributeError, TypeError):
        return ""


def _norm(s):
    """去掉字符串内全部空白（含全角空格），用于号牌/编号/日期/数值等字段。"""
    return re.sub(r'[\s\u3000]+', '', s or '')


def _find_table_by_keyword(doc, keywords: Iterable[str]) -> Optional[Table]:
    """按内容定位表格：返回第一张包含任一关键词的表格（不依赖表格序号）。"""
    for table in doc.tables:
        cells, _ = _table_grid(table)
        for cell in cells:
            t = cell.text
            for kw in keywords:
                if kw in t:
                    return table
    return None


_DATE_RE = re.compile(r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日')


def _fmt_date(m: Match[str]) -> str:
    return "%d年%02d月%02d日" % (int(m.group(1)), int(m.group(2)), int(m.group(3)))


def _extract_signature_dates(doc) -> Tuple[str, str]:
    """提取“检验”“审核”签名日期（兼容日期与标签同格/分格两种版式）。"""
    jy = sh = ""
    for table in doc.tables:
        cells, ncol = _table_grid(table)
        nrow = len(table.rows)
        for r in range(nrow):
            comp = [_norm(cells[r * ncol + c].text) for c in range(ncol)]
            for i in range(ncol):
                t = comp[i]
                if not jy:
                    m = re.search(r'检验\s*[:：]?\s*(\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)', t)
                    if m:
                        dm = _DATE_RE.search(m.group(1))
                        if dm:
                            jy = _fmt_date(dm)
                    else:
                        m2 = _DATE_RE.fullmatch(t)
                        if m2:
                            left = "".join(comp[max(0, i - 3):i])
                            if re.search(r'检验[:：]?$', left) and '审核' not in left and '校核' not in left:
                                jy = _fmt_date(m2)
                if not sh:
                    m = re.search(r'(?:审核|校核|审\s*核)\s*[:：]?\s*(\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)', t)
                    if m:
                        dm = _DATE_RE.search(m.group(1))
                        if dm:
                            sh = _fmt_date(dm)
                    else:
                        m2 = _DATE_RE.fullmatch(t)
                        if m2:
                            left = "".join(comp[max(0, i - 3):i])
                            if re.search(r'(?:审核|校核)[:：]?$', left):
                                sh = _fmt_date(m2)
    return jy, sh


def _extract_medium_list(table: Table) -> List[Dict[str, str]]:
    """按表头列名提取适装介质列表（兼容 9/10 列两种版式）。"""
    cells, ncol = _table_grid(table)
    nrow = len(table.rows)
    hdr_r = None
    for r in range(nrow):
        row_comp = [_norm(cells[r * ncol + c].text) for c in range(ncol)]
        if any(x == "序号" for x in row_comp) and any("介质名称" in x for x in row_comp):
            hdr_r = r
            break
    if hdr_r is None:
        return []
    hdr = [_norm(cells[hdr_r * ncol + c].text) for c in range(ncol)]

    def col(pred):
        for c, x in enumerate(hdr):
            if pred(x):
                return c
        return -1

    c_seq = col(lambda x: x == "序号")
    c_name = col(lambda x: x.startswith("介质名称"))
    c_un = col(lambda x: x.startswith("UN"))
    c_cat = col(lambda x: x.startswith("类别"))
    c_pack = col(lambda x: x.startswith("包装"))
    c_note = col(lambda x: x == "备注")

    def gv(r, c):
        # 列名缺失时 col() 返回 -1，负索引会静默绕回上一行末格取到错值，必须拦为空
        return "" if c < 0 else _cell_text((cells, ncol), r, c)

    out = []
    for r in range(hdr_r + 1, nrow):
        seq = gv(r, c_seq)
        if not seq.isdigit():
            continue
        name = gv(r, c_name)
        if name in ("", "——"):
            continue
        out.append({
            "序号": seq,
            "介质名称": name,
            "UN号": gv(r, c_un),
            "类别及项别": gv(r, c_cat),
            "包装类别": gv(r, c_pack),
            "备注": gv(r, c_note),
        })
    return out


def _extract_thickness(table: Table) -> Dict[str, str]:
    """按“点数→数值”解析测厚表（封头/筒体都可，兼容 18/19 列）。"""
    cells, ncol = _table_grid(table)
    nrow = len(table.rows)
    data = {}
    int_re = re.compile(r'^\d+$')
    dec_re = re.compile(r'^\d+\.\d+$')
    for r in range(nrow):
        section = None
        pending = None
        for c in range(ncol):
            t = _norm(cells[r * ncol + c].text)
            if t == "":
                continue
            if "封头" in t:
                section = "封头"
                pending = None
            elif "筒体" in t:
                section = "筒体"
                pending = None
            elif int_re.match(t) and section is not None:
                pending = int(t)
            elif dec_re.match(t) and section is not None and pending is not None:
                key = "%s%d" % (section, pending)
                if key not in data:
                    data[key] = t
    return data


_APPENDIX_HEADING = "充装介质附页"
_APPENDIX_MARK = "充装以下介质"


def _appendix_boundary(body_kids) -> Optional[Tuple[int, int]]:
    """返回附页节在 body 直接子元素中的 (起点, 终点)。

    起点：含“充装以下介质”标记的整表；再向前回溯到“…充装介质附页”标题段，
    使整页（标题/报告编号/整表/注释/空行）完整。终点：body 末尾 sectPr 之前。
    找不到整表时返回 None。
    """
    ti = None
    for i, kid in enumerate(body_kids):
        if kid.tag == qn('w:tbl') and _APPENDIX_MARK in "".join(
                t.text or "" for t in kid.iter(qn('w:t'))):
            ti = i
    if ti is None:
        return None
    start = ti
    for i in range(ti - 1, -1, -1):
        kid = body_kids[i]
        if kid.tag != qn('w:p'):
            break
        if _APPENDIX_HEADING in "".join(t.text or "" for t in kid.iter(qn('w:t'))):
            start = i
            break
    end = len(body_kids)
    if body_kids and body_kids[-1].tag == qn('w:sectPr'):
        end -= 1
    return start, end


def _collect_style_chain(src_styles_el, style_ids) -> List[Any]:
    """按 styleId 收集样式定义及其 basedOn 依赖链（深拷贝）。"""
    by_id = {s.get(qn('w:styleId')): s for s in src_styles_el.findall(qn('w:style'))}
    picked, queue = [], list(style_ids)
    seen = set()
    while queue:
        sid = queue.pop()
        if not sid or sid in seen or sid not in by_id:
            continue
        seen.add(sid)
        st = by_id[sid]
        picked.append(copy.deepcopy(st))
        based = st.find(qn('w:basedOn'))
        if based is not None:
            queue.append(based.get(qn('w:val')))
    return picked


# noinspection PyProtectedMember
def _effective_refs(doc) -> List[Tuple[str, str, Optional[str]]]:
    """按文档节顺序推导最后一节“实际生效”的页眉页脚引用。

    OOXML 中某节没有显式 headerReference/footerReference 时继承上一节。
    返回 [(kind('header'/'footer'), w:type, rId), ...]，按 header 在前、
    footer 在后的顺序排列（符合 sectPr schema 顺序）。
    """
    sect_els = [p.find(qn('w:pPr') + '/' + qn('w:sectPr'))
                for p in doc.element.body.iter(qn('w:p'))]
    sect_els = [s for s in sect_els if s is not None]
    sect_els.append(doc.element.body.find(qn('w:sectPr')))

    refs = {}  # (kind, w:type) -> rid
    for sect in sect_els:
        for kind, tag in (('header', 'w:headerReference'),
                          ('footer', 'w:footerReference')):
            for ref in sect.findall(qn(tag)):
                wtype = ref.get(qn('w:type')) or 'default'
                refs[(kind, wtype)] = ref.get(qn('r:id'))

    ordered = sorted(refs.items(), key=lambda kv: (0 if kv[0][0] == 'header' else 1,
                                                   kv[0][1]))
    return [(kind, wtype, rid) for (kind, wtype), rid in ordered]


# 页码域识别：仅匹配域指令的第一个域名为 PAGE（不能误匹配 NUMPAGES）
_PG_FIELD_NAME_RE = re.compile(r'^\s*PAGE\b', re.IGNORECASE)


def _strip_page_number_fields(xml_root) -> int:
    """就地删除页眉/页脚 XML 中全部 PAGE 页码域，返回删除元素数。

    兼容两种域形态：w:fldSimple 简单域；fldChar begin/separate/end 复合域
    （页脚里 mc:Choice 与 mc:Fallback 各有一份，都会被清掉）。只删域 run，
    其余文字/段落/浮动文本框骨架原样保留。check.py 中有同名镜像实现，
    修改口径时两边必须同步。
    """
    n = 0
    for fs in list(xml_root.iter(qn('w:fldSimple'))):
        if _PG_FIELD_NAME_RE.match(fs.get(qn('w:instr')) or ''):
            parent = fs.getparent()
            if parent is not None:
                parent.remove(fs)
                n += 1
    for para in list(xml_root.iter(qn('w:p'))):
        stack = []  # [{'instr': 已收集指令文本, 'runs': [域内 run,...]}]
        for child in list(para):
            if child.tag != qn('w:r'):
                continue
            top = stack[-1] if stack else None
            if top is not None:
                top['runs'].append(child)
                for it in child.findall(qn('w:instrText')):
                    top['instr'] += it.text or ''
            fld = child.find(qn('w:fldChar'))
            if fld is None:
                continue
            ctype = fld.get(qn('w:fldCharType'))
            if ctype == 'begin':
                stack.append({'instr': '', 'runs': [child]})
            elif ctype == 'end' and stack:
                entry = stack.pop()
                if _PG_FIELD_NAME_RE.match(entry['instr']):
                    for r in entry['runs']:
                        parent = r.getparent()
                        if parent is not None:
                            parent.remove(r)
                            n += 1
    return n


# noinspection PyProtectedMember
def _extract_appendix(doc, page_start=None):
    """无损提取报告末尾的“充装介质附页”整节。

    为做到与原页逐元素一致，深拷贝四类对象：
      1) 附页节的全部 body 元素（标题段、报告编号段、整张带框表格、注释、空行）；
      2) 附页节的 sectPr（页面尺寸/页边距等），并把继承来的页眉页脚全部改成
         显式引用（源附页节本身只有页脚引用，页眉继承自上一节）；
      3) 引用到的页眉页脚部件（页脚剔除 PAGE 页码域后装入成品包，其余原样）；
      4) 内容与页眉页脚引用到的样式定义（含 basedOn 依赖链）。
    page_start：附页在源报告该节中的节内页码，写入 pgNumType 的 start，
    使成品附页页码与源页显示同一数字（PAGE 域始终由排版引擎实时计算）。

    返回 dict：elems / text / sectpr / parts / styles；无附页时返回 None。
    """
    # 快捷预检：正文完全不含附页标记时直接返回（大多数报告的常态路径）
    _body_full = "".join(
        t.text or "" for t in doc.element.body.iter(qn('w:t')))
    if _APPENDIX_MARK not in _body_full and _APPENDIX_HEADING not in _body_full:
        return None

    kids = list(doc.element.body)
    boundary = _appendix_boundary(kids)
    if boundary is None:
        return None
    start, end = boundary

    elems = [copy.deepcopy(kids[i]) for i in range(start, end)]
    text = "\n".join(
        "".join(t.text or "" for t in e.iter(qn('w:t'))) for e in elems)

    sectpr = copy.deepcopy(doc.element.body.find(qn('w:sectPr')))
    for tag in ('w:headerReference', 'w:footerReference'):
        for ref in sectpr.findall(qn(tag)):
            sectpr.remove(ref)

    # 生效的页眉页脚引用（含继承），在 sectPr 顶部按 schema 顺序写为显式引用
    eff = _effective_refs(doc)
    parts = []
    doc_part = doc.part
    for kind, wtype, rid in eff:
        if not rid or rid not in doc_part.rels:
            continue
        ref = OxmlElement('w:' + kind + 'Reference')
        ref.set(qn('w:type'), wtype)
        ref.set(qn('r:id'), rid)
        sectpr.insert(len(parts), ref)
        p = doc_part.rels[rid].target_part
        part_blob = p.blob
        if kind == 'footer':
            # 应需求取消附页页码：装入前剔除页脚中的 PAGE 域，其余原样
            try:
                froot = etree.fromstring(part_blob)
                _strip_page_number_fields(froot)
                part_blob = etree.tostring(
                    froot, xml_declaration=True, encoding='UTF-8', standalone=True)
            except (etree.XMLSyntaxError, ValueError):
                part_blob = p.blob
        # 页眉页脚自身的内部关系（如图片）也一并带走：
        # (旧rId, 关系类型, 扩展名, 内容类型, blob)
        sub = []
        for rr in p.rels.values():
            if rr.is_external:
                continue
            tp = rr.target_part
            if tp is p:
                continue
            ext = str(tp.partname).rsplit('.', 1)[-1]
            sub.append((rr.rId, rr.reltype, ext, tp.content_type, tp.blob))
        parts.append((rid, kind, wtype, p.content_type, part_blob, sub))

    # 附页正文引用的内嵌部件（印章/图片等 blip/VML/OLE）
    body_media = []
    media_seen = set()
    for e in elems:
        for el in e.iter():
            for attr in ('embed', 'id', 'link'):
                mid = el.get(qn('r:' + attr))
                if not mid or mid in media_seen or mid not in doc_part.rels:
                    continue
                rel = doc_part.rels[mid]
                if rel.is_external:
                    continue
                tp = rel.target_part
                media_seen.add(mid)
                ext = str(tp.partname).rsplit('.', 1)[-1]
                body_media.append((mid, rel.reltype, ext, tp.content_type, tp.blob))

    # 页码起始值仍按源页写入（附页 PAGE 域已剔除，当前无可见效果；
    # 保留是为日后若恢复页码，整套机制可直接生效）
    if page_start:
        pgn = sectpr.find(qn('w:pgNumType'))
        if pgn is None:
            pgn = OxmlElement('w:pgNumType')
            anchor = sectpr.find(qn('w:pgMar'))
            if anchor is not None:
                anchor.addnext(pgn)
            else:
                sectpr.insert(0, pgn)
        pgn.set(qn('w:start'), str(int(page_start)))

    # 收集引用到的样式：附页内容的 tblStyle/pStyle/rStyle + 页眉页脚内的 pStyle/rStyle
    style_ids = set()
    for e in elems:
        for tag in ('w:tblStyle', 'w:pStyle', 'w:rStyle'):
            style_ids.update(x.get(qn('w:val')) for x in e.iter(qn(tag)))
    for part_info in parts:
        try:
            xml_root = etree.fromstring(part_info[4])
            for tag in ('w:pStyle', 'w:rStyle'):
                style_ids.update(x.get(qn('w:val')) for x in xml_root.iter(qn(tag)))
        except (etree.XMLSyntaxError, ValueError, AttributeError):
            pass
    styles = _collect_style_chain(doc.styles.element, style_ids)

    # 源报告是否关闭标点数距配对（影响“（）、：”等标点在整行中的亚像素位置）
    npk = doc.settings.element.find(qn('w:noPunctuationKerning')) is not None

    return {'elems': elems, 'text': text, 'sectpr': sectpr,
            'parts': parts, 'styles': styles, 'npk': npk,
            'body_media': body_media}


def _cell_runs(grid: Tuple[Any, int], row: int, col: int) -> List[Dict[str, Any]]:
    """按 run 提取单元格内容（用于分仓数量，保留下划线等格式）。"""
    cells, ncol = grid
    try:
        cell = cells[row * ncol + col]
    except (IndexError, TypeError):
        return []
    runs = []
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            runs.append({
                'text': run.text or '',
                'underline': bool(run.font.underline) if run.font is not None else False,
            })
    return runs


def set_cell_center(grid: Tuple[Any, int], row: int, col: int, value: Any) -> None:
    try:
        cells, col_count = grid
        cell = cells[row * col_count + col]
        cell.text = ""
        run = cell.paragraphs[0].add_run(str(value))
        _set_run_font(run)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    except (IndexError, AttributeError, TypeError, ValueError):
        pass



# ============================================================================
# 【复制代码】以下函数用于从源文档读取/复制数据
# ============================================================================

def set_cell_with_runs(grid: Tuple[Any, int], row: int, col: int,
                       runs_info: List[Dict[str, Any]]) -> bool:
    """根据run信息列表设置单元格内容，保留下划线格式"""
    try:
        cells, col_count = grid
        cell = cells[row * col_count + col]
        cell.text = ""

        for run_info in runs_info:
            run = cell.paragraphs[0].add_run(run_info['text'])
            _set_run_font(run)
            if run_info.get('underline'):
                run.font.underline = True

        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        return True
    except (IndexError, AttributeError, TypeError, ValueError) as e:
        print(f"  设置单元格内容失败: {e}")
        return False


# ============================================================================
# 【复制代码】从检验报告中提取数据
# ============================================================================

class ReportExtractError(Exception):
    """报告提取致命错误：文件损坏或关键标识缺失。

    fail-closed 原则：出现此错误时调用方必须标记失败且不得生成成品，
    避免“空数据→全空原始记录→success=True”的静默失败。
    """


# noinspection PyProtectedMember
def extract_data_from_report(report_path: str, appendix_page: Optional[int] = None) -> Dict[str, Any]:
    """从检验报告中提取数据。致命错误抛 ReportExtractError，不返回空数据。"""
    data = {
        "机动车号牌": "", "使用单位": "", "道路运输证号": "", "单位地址": "",
        "总质量": "", "核定载质量": "", "制造企业": "", "制造日期": "",
        "设计代码": "", "产品标准": "", "产品型号": "", "VIN码": "",
        "罐体编号": "", "罐体容积": "", "罐体外形尺寸": "", "分仓数量": "",
        "封头材质": "", "筒体材质": "", "封头厚度": "", "筒体厚度": "",
        "装运介质": "", "报告编号": "",
        "封头1": "", "封头2": "", "封头3": "", "封头4": "",
        "封头17": "", "封头18": "", "封头19": "", "封头20": "",
        "筒体5": "", "筒体6": "", "筒体7": "", "筒体8": "", "筒体9": "", "筒体10": "",
        "筒体11": "", "筒体12": "", "筒体13": "", "筒体14": "", "筒体15": "", "筒体16": "",
        "检验日期": "", "审核日期": "", "下次检验日期": ""
    }

    try:
        # 【文件操作代码】打开检验报告文档进行读取
        doc = Document(report_path)

        # 提取报告编号
        for para in doc.paragraphs:
            text = para.text.strip()
            if "报告编号" in text:
                match = re.search(r'报告编号[：:]?\s*(\S+)', text)
                if match:
                    num = match.group(1)
                    if len(num) > len(data["报告编号"]):
                        data["报告编号"] = num

        # 基本信息表（机动车号牌/使用单位/道路运输证号/单位地址）
        basic = _find_table_by_keyword(doc, ["道路运输证号"])
        if basic is not None:
            g = _table_grid(basic)
            data["机动车号牌"] = _norm(_cell_text(g, 0, 1))
            data["使用单位"] = _cell_text(g, 2, 1)
            data["道路运输证号"] = _norm(_cell_text(g, 2, 4))
            data["单位地址"] = _cell_text(g, 3, 1)

        # 罐体基本资料表（制造/材质/厚度/装运介质/分仓数量 等）
        base = _find_table_by_keyword(doc, ["罐体材质", "分仓数量"])
        if base is not None:
            g = _table_grid(base)
            data["罐体编号"] = _norm(_cell_text(g, 0, 5))
            data["制造日期"] = _norm(_cell_text(g, 1, 1))
            data["VIN码"] = _norm(_cell_text(g, 1, 5))
            data["制造企业"] = _cell_text(g, 2, 1)
            data["设计代码"] = _norm(_cell_text(g, 3, 1))
            data["产品标准"] = _norm(_cell_text(g, 3, 5))
            data["产品型号"] = _norm(_cell_text(g, 4, 1))
            data["罐体容积"] = _norm(_cell_text(g, 4, 5))
            data["罐体外形尺寸"] = _norm(_cell_text(g, 5, 1))
            data["分仓数量"] = _cell_runs(g, 5, 5)
            data["总质量"] = _norm(_cell_text(g, 6, 1))
            data["核定载质量"] = _norm(_cell_text(g, 6, 5))
            data["封头材质"] = _cell_text(g, 8, 2)
            data["筒体材质"] = _cell_text(g, 9, 2)
            data["封头厚度"] = _norm(_cell_text(g, 8, 6))
            data["筒体厚度"] = _norm(_cell_text(g, 9, 6))
            data["装运介质"] = _cell_text(g, 10, 1)

        # 适装介质列表
        medium_table = _find_table_by_keyword(doc, ["介质名称"])
        if medium_table is not None:
            data["适装介质列表"] = _extract_medium_list(medium_table)

        # 测厚记录
        thickness_table = _find_table_by_keyword(doc, ["测厚记录", "测厚点部位图"])
        if thickness_table is not None:
            data.update(_extract_thickness(thickness_table))

        # 检验/审核签名日期
        data["检验日期"], data["审核日期"] = _extract_signature_dates(doc)

        # 下次检验日期（多个命中时取第一个带日期的，与 check.py 口径一致）
        _nxt = ""
        for table in doc.tables:
            cells, col_count = _table_grid(table)
            for r in range(len(table.rows)):
                row_base = r * col_count
                for c in range(col_count):
                    t = _norm(cells[row_base + c].text)
                    if "下次检验日期" in t:
                        match = re.search(r'下次检验日期[：:]?\s*(\d{4}\s*年\s*\d{1,2}\s*月)', t)
                        if match:
                            _nxt = _norm(match.group(1))
                            break
                if _nxt:
                    break
            if _nxt:
                break
        data["下次检验日期"] = _nxt

        # 充装介质附页：无条件探测正文是否含附页（兼容装运栏各种措辞变体），
        # 装运栏明确写了触发词但找不到附页时标记缺失（汇总名单交人工核查）
        zy = _norm(data.get("装运介质", ""))
        claims_appendix = ("附" in zy) or bool(re.search(r"末页|后页|清单", zy))
        appendix = _extract_appendix(doc, appendix_page)
        data["_附页缺失"] = bool(claims_appendix and appendix is None)
        if appendix:
            data["附页文字"] = appendix["text"]
            data["_附页"] = appendix

        # fail-closed：关键标识缺失说明文件损坏/异版式，不得生成空白成品
        if not data["报告编号"] or not data["机动车号牌"]:
            raise ReportExtractError(
                "关键标识字段缺失（报告编号=%r，号牌=%r），疑似文件损坏或异版式"
                % (data["报告编号"], data["机动车号牌"]))

    except ReportExtractError:
        raise
    # fail-closed 边界：任何意外异常都要转成致命错误而不是产出空成品
    # noinspection PyBroadException
    except Exception as e:
        print(f"  提取数据失败: {e}")
        import traceback
        traceback.print_exc()
        raise ReportExtractError(f"提取数据失败: {e}")

    return data


# ============================================================================
# 【粘贴代码】填充模板文档的核心函数 - 将所有数据写入模板
# ============================================================================

# noinspection PyProtectedMember
def _append_appendix_section(doc, appendix) -> None:
    """把报告附页整节无损接到成品末尾：样式、页眉页脚、分节几何全部照搬。"""
    body = doc.element.body

    # 1) 样式合并：缺 id 直接加入；同 id 但定义不同（如模板的“5/6”与源报告不同）
    #    则给源样式改一个唯一 id 注册，并把拷贝内容里的引用全部重定向。
    styles_el = doc.styles.element
    tpl_styles = {}
    for s in styles_el.findall(qn('w:style')):
        tpl_styles[s.get(qn('w:styleId'))] = etree.tostring(s, method='c14n')

    by_src = {}
    for st in appendix['styles']:
        by_src.setdefault(st.get(qn('w:styleId')), st)

    # 第一遍：确定每个源样式在成品中使用的 id（碰撞时改名）
    style_id_map = {}      # 源 id -> 成品 id
    used_ids = set(tpl_styles)
    for sid, st in by_src.items():
        sig = etree.tostring(st, method='c14n')
        if sid in tpl_styles and tpl_styles[sid] == sig:
            style_id_map[sid] = sid
            continue
        nid = sid if sid not in tpl_styles else 'apd_' + sid
        i = 1
        while nid in used_ids:
            nid = 'apd_%s_%d' % (sid, i)
            i += 1
        style_id_map[sid] = nid
        used_ids.add(nid)

    # 第二遍：注册新样式（basedOn/next/link 一并指向新 id）
    for sid, st in by_src.items():
        nid = style_id_map[sid]
        if sid in tpl_styles and nid == sid:
            continue
        new_st = copy.deepcopy(st)
        new_st.set(qn('w:styleId'), nid)
        for link_tag in ('w:basedOn', 'w:next', 'w:link'):
            for le in new_st.findall(qn(link_tag)):
                ref_id = le.get(qn('w:val'))
                if ref_id in style_id_map:
                    le.set(qn('w:val'), style_id_map[ref_id])
        styles_el.append(new_st)

    def _remap_style_refs(kid):
        for tag in ('w:pStyle', 'w:rStyle', 'w:tblStyle'):
            for style_el in kid.iter(qn(tag)):
                ref_id = style_el.get(qn('w:val'))
                if ref_id in style_id_map and style_id_map[ref_id] != ref_id:
                    style_el.set(qn('w:val'), style_id_map[ref_id])

    for el_root in appendix['elems']:
        _remap_style_refs(el_root)

    # 2) 页眉页脚部件原样装入成品包（含其引用的图片），重映射关系 ID
    rid_map = {}
    pkg = doc.part.package

    def _add_media_part(owner, reltype, ext, ctype, data_blob):
        """把一个内嵌媒体部件装入包并挂到 owner 部件下，返回新关系 ID。"""
        name_tpl = '/word/media/appendix_%%d.%s' % (ext or 'bin')
        media = Part(pkg.next_partname(name_tpl), ctype, data_blob, pkg)
        return owner.relate_to(media, reltype)

    for old_rid, kind, _wtype, content_type, blob, sub in appendix['parts']:
        name_tpl = '/word/footer%d.xml' if kind == 'footer' else '/word/header%d.xml'
        part = Part(pkg.next_partname(name_tpl), content_type, blob, pkg)
        new_rid = doc.part.relate_to(
            part, RT.FOOTER if kind == 'footer' else RT.HEADER)
        rid_map[old_rid] = new_rid
        # 页眉页脚内的图片：挂到“新页眉页脚部件”自身；样式引用按碰撞映射改名
        sub_map = {}
        for srid, sreltype, sext, sct, sblob in sub:
            sub_map[srid] = _add_media_part(part, sreltype, sext, sct, sblob)
        if sub_map or any(v != k for k, v in style_id_map.items()):
            hf_xml = etree.fromstring(part.blob)
            _remap_style_refs(hf_xml)
            for node in hf_xml.iter():
                for attr in ('embed', 'id', 'link'):
                    old = node.get(qn('r:' + attr))
                    if old in sub_map:
                        node.set(qn('r:' + attr), sub_map[old])
            part._blob = etree.tostring(
                hf_xml, xml_declaration=True, encoding='UTF-8', standalone=True)

    # 3) 附页正文图片（印章等）：挂到文档主部件，重映射拷贝元素中的 rId
    media_map = {}
    for mid, mreltype, mext, mct, mblob in appendix.get('body_media', []):
        media_map[mid] = _add_media_part(doc.part, mreltype, mext, mct, mblob)
    for el_root in appendix['elems']:
        for node in el_root.iter():
            for attr in ('embed', 'id', 'link'):
                old = node.get(qn('r:' + attr))
                if old in media_map:
                    node.set(qn('r:' + attr), media_map[old])

    # 4) 复制源节属性，改写其中的页眉页脚关系 ID
    src_sect = copy.deepcopy(appendix['sectpr'])
    for ref in src_sect.findall(qn('w:headerReference')) + \
            src_sect.findall(qn('w:footerReference')):
        old = ref.get(qn('r:id'))
        if old in rid_map:
            ref.set(qn('r:id'), rid_map[old])

    # 4) 同步源报告的 noPunctuationKerning（否则标点压缩差异会让正文行错位）
    if appendix.get('npk'):
        st = doc.settings.element
        if st.find(qn('w:noPunctuationKerning')) is None:
            npk = OxmlElement('w:noPunctuationKerning')
            anchor = st.find(qn('w:characterSpacingControl'))
            if anchor is not None:
                anchor.addprevious(npk)
            else:
                st.append(npk)

    # 5) 原模板节在此结束：用一个带 sectPr 的空段落做（下一页）分节符，
    #    附页元素随后整体追加，body 末尾 sectPr 换成源附页节的属性。
    tmpl_sect = body.find(qn('w:sectPr'))
    if tmpl_sect is not None:
        body.remove(tmpl_sect)
    break_p = OxmlElement('w:p')
    if tmpl_sect is not None:
        ppr = OxmlElement('w:pPr')
        ppr.append(tmpl_sect)
        break_p.append(ppr)
    body.append(break_p)
    for elem in appendix['elems']:
        body.append(elem)
    body.append(src_sect)


# 日期标签格：只允许“检验/校核(审核)：”后跟空白/下划线/年月日/数字，
# 精确锚定日期格，避免误伤“检验结论/审核意见”等其它单元格。
_DATE_CELL_JY_RE = re.compile(r"^检验\s*[:：]?[\s_年月日0-9]*$")
_DATE_CELL_SH_RE = re.compile(r"^(?:校核|审核)\s*[:：]?[\s_年月日0-9]*$")


def _fill_date_cell(cell, label_re, value):
    """填充检验/审核日期：有日期则 run 内替换；空占位则直接追加完整日期。"""
    t = cell.text.strip()
    if not value or not label_re.match(t):
        return
    existing = re.search(r"\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日", t)
    if existing:
        old = existing.group(0)
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                if run.text and old in run.text:
                    run.text = run.text.replace(old, value)
                if run.font:
                    _set_run_font(run)
    else:
        # 空占位（无数字日期）：按既定策略直接追加完整日期
        run = cell.paragraphs[0].add_run(value)
        _set_run_font(run)


def _fill_next_date_cell(cell, value):
    """填充下次检验日期（年月）：有日期则替换；空占位则追加。"""
    t = cell.text.strip()
    if "下次检验日期" not in t or not value:
        return
    existing = re.search(r"\d{4}\s*年\s*\d{1,2}\s*月", t)
    if existing:
        old = existing.group(0)
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                if run.text and old in run.text:
                    run.text = run.text.replace(old, value)
                if run.font:
                    _set_run_font(run)
    else:
        run = cell.paragraphs[0].add_run(value)
        _set_run_font(run)


def fill_template(input_data: Dict[str, Any], output_path: str) -> None:
    # 【文件操作代码】打开模板文档
    doc = Document(TEMPLATE_FILE)

    for para in doc.paragraphs:
        if "记录编号：" in para.text:
            para.text = re.sub(r'记录编号：\S*', "记录编号：%s" % input_data.get("报告编号", ""), para.text)
            for run in para.runs:
                if run.font:
                    _set_run_font(run)

    table0 = doc.tables[0]
    g0 = _table_grid(table0)
    set_cell_center(g0, 0, 2, input_data.get("机动车号牌", ""))
    set_cell_center(g0, 2, 2, input_data.get("使用单位", ""))
    set_cell_center(g0, 3, 2, input_data.get("单位地址", ""))
    set_cell_center(g0, 2, 6, input_data.get("道路运输证号", ""))
    set_cell_center(g0, 5, 2, input_data.get("总质量", ""))
    set_cell_center(g0, 5, 6, input_data.get("核定载质量", ""))
    set_cell_center(g0, 6, 2, input_data.get("制造企业", ""))
    set_cell_center(g0, 6, 6, input_data.get("制造日期", ""))
    set_cell_center(g0, 7, 2, input_data.get("设计代码", ""))
    set_cell_center(g0, 7, 6, input_data.get("产品标准", ""))
    set_cell_center(g0, 8, 2, input_data.get("产品型号", ""))
    set_cell_center(g0, 8, 6, input_data.get("VIN码", ""))
    set_cell_center(g0, 9, 2, input_data.get("罐体编号", ""))
    set_cell_center(g0, 9, 6, input_data.get("罐体容积", ""))
    set_cell_center(g0, 10, 2, input_data.get("罐体外形尺寸", ""))

    fen_cang_data = input_data.get("分仓数量", "")
    if isinstance(fen_cang_data, list):
        # 空单元格提取结果是空 run 列表，绝不能把 "[] 仓" 写进成品
        if fen_cang_data and "".join(r.get('text', '') for r in fen_cang_data).strip():
            set_cell_with_runs(g0, 10, 6, fen_cang_data)
        else:
            set_cell_center(g0, 10, 6, "")
    else:
        _fc = str(fen_cang_data).strip()
        set_cell_center(g0, 10, 6, (_fc + " 仓") if _fc else "")

    set_cell_center(g0, 12, 3, input_data.get("封头材质", ""))
    set_cell_center(g0, 13, 3, input_data.get("筒体材质", ""))
    set_cell_center(g0, 12, 7, input_data.get("封头厚度", ""))
    set_cell_center(g0, 13, 7, input_data.get("筒体厚度", ""))
    set_cell_center(g0, 14, 2, input_data.get("装运介质", ""))

    # 填充适装介质列表（table0 row17 起；模板固定 13 行，超出时克隆最后一行动态扩行）
    medium_list = input_data.get("适装介质列表", [])
    media_start_row = 17
    media_capacity = len(table0.rows) - media_start_row
    extra_count = len(medium_list) - media_capacity
    if extra_count > 0:
        # 介质区位于 table0 末尾，直接克隆最后一个介质行（保留边框/合并/列宽等格式）
        # noinspection PyProtectedMember
        last_media_tr = table0.rows[len(table0.rows) - 1]._tr
        for _ in range(extra_count):
            table0._tbl.append(copy.deepcopy(last_media_tr))
        g0 = _table_grid(table0)  # 行结构已变，重新物化网格
    for i, row_data in enumerate(medium_list):
        target_row = media_start_row + i
        if target_row >= len(table0.rows):
            break
        set_cell_center(g0, target_row, 0, row_data.get("序号", ""))
        set_cell_center(g0, target_row, 1, row_data.get("介质名称", ""))
        set_cell_center(g0, target_row, 4, row_data.get("UN号", ""))
        set_cell_center(g0, target_row, 5, row_data.get("类别及项别", ""))
        set_cell_center(g0, target_row, 6, row_data.get("包装类别", ""))
        set_cell_center(g0, target_row, 8, row_data.get("备注", ""))

    table2 = doc.tables[3]
    g3 = _table_grid(table2)
    set_cell_center(g3, 3, 3, input_data.get("封头1", ""))
    set_cell_center(g3, 4, 3, input_data.get("封头2", ""))
    set_cell_center(g3, 5, 3, input_data.get("封头3", ""))
    set_cell_center(g3, 6, 3, input_data.get("封头4", ""))
    set_cell_center(g3, 3, 5, input_data.get("封头17", ""))
    set_cell_center(g3, 4, 5, input_data.get("封头18", ""))
    set_cell_center(g3, 5, 5, input_data.get("封头19", ""))
    set_cell_center(g3, 6, 5, input_data.get("封头20", ""))

    set_cell_center(g3, 3, 10, input_data.get("筒体5", ""))
    set_cell_center(g3, 4, 10, input_data.get("筒体6", ""))
    set_cell_center(g3, 5, 10, input_data.get("筒体7", ""))
    set_cell_center(g3, 6, 10, input_data.get("筒体8", ""))
    set_cell_center(g3, 3, 13, input_data.get("筒体9", ""))
    set_cell_center(g3, 4, 13, input_data.get("筒体10", ""))
    set_cell_center(g3, 5, 13, input_data.get("筒体11", ""))
    set_cell_center(g3, 6, 13, input_data.get("筒体12", ""))
    set_cell_center(g3, 3, 16, input_data.get("筒体13", ""))
    set_cell_center(g3, 4, 16, input_data.get("筒体14", ""))
    set_cell_center(g3, 5, 16, input_data.get("筒体15", ""))
    set_cell_center(g3, 6, 16, input_data.get("筒体16", ""))

    # 检验/审核日期（有日期替换，空占位追加）
    cells3, col3 = g3
    for r in range(len(table2.rows)):
        row_base = r * col3
        for c in range(col3):
            cell = cells3[row_base + c]
            _fill_date_cell(cell, _DATE_CELL_JY_RE,
                            input_data.get("检验日期", ""))
            _fill_date_cell(cell, _DATE_CELL_SH_RE,
                            input_data.get("审核日期", ""))

    # 下次检验日期
    for table in doc.tables:
        cells, col_count = _table_grid(table)
        for r in range(len(table.rows)):
            row_base = r * col_count
            for c in range(col_count):
                try:
                    _fill_next_date_cell(
                        cells[row_base + c],
                        input_data.get("下次检验日期", ""))
                except (IndexError, AttributeError, TypeError, ValueError):
                    pass

    # 【附页代码】装运介质为“见附页/附表”时，把报告末尾附页整节无损搬入
    appendix = input_data.get("_附页")
    if appendix and appendix.get("elems"):
        _append_appendix_section(doc, appendix)

    # 【文件操作代码】原子保存：先写同目录临时文件，写完再替换正式文件，
    # 避免保存中途异常/被杀导致正式成品损坏（旧成品也不会被截断）
    tmp_out = "%s.%d.tmp" % (output_path, os.getpid())
    try:
        doc.save(tmp_out)
        os.replace(tmp_out, output_path)
    finally:
        if os.path.exists(tmp_out):
            try:
                os.remove(tmp_out)
            except OSError:
                pass


def process_single_report(args: Tuple[str, str, str, Optional[int]]) -> Dict[str, Any]:
    """处理单个报告文件（用于多进程并行）"""
    filename, reports_folder, output_folder, appendix_page = args

    # 【文件操作代码】拼接报告文件路径
    report_path = os.path.join(reports_folder, filename)

    try:
        data = extract_data_from_report(report_path, appendix_page)

        # 成品文件名 = 报告编号 + 机动车号牌 + 原始记录（两字段已经过 fail-closed 校验，必非空）
        # 剔除 Windows 文件名非法字符，防止异常数据导致写盘失败
        safe_name = re.sub(r'[\\/:*?"<>|]', '', data["报告编号"] + data["机动车号牌"])
        # filename 可能含嵌套子目录：保留源的目录结构，只替换文件名部分
        sub_dir = os.path.dirname(filename)
        output_filename = (os.path.join(sub_dir, safe_name + '原始记录.docx')
                           if sub_dir else safe_name + '原始记录.docx')
        # 【文件操作代码】拼接输出文件路径（filename 可能含嵌套子目录）
        output_path = os.path.join(output_folder, output_filename)
        out_parent = os.path.dirname(output_path)
        if out_parent and not os.path.exists(out_parent):
            os.makedirs(out_parent, exist_ok=True)

        fill_template(data, output_path)

        zy_text = data.get("装运介质", "").strip()
        special_media = zy_text and ("汽油" not in zy_text and "柴油" not in zy_text)
        has_appendix = bool(data.get("_附页"))
        appendix_missing = bool(data.get("_附页缺失"))

        plate_number = data.get("机动车号牌", "").strip()
        guangxi_plate = plate_number.startswith("桂")

        # 单位地址超长提醒：模板行高固定为 exact，超长地址换行后 PDF 可能截断
        addr = data.get("单位地址", "") or ""
        addr_long = len(addr) > 20 or "\n" in addr or "\r" in addr

        return {
            'success': True,
            'filename': filename,
            'output_filename': output_filename,
            'output_path': output_path,
            'rel_output': os.path.relpath(output_path, OUTPUT_FOLDER),
            'special_media': special_media,
            'guangxi_plate': guangxi_plate,
            'has_appendix': has_appendix,
            'appendix_missing': appendix_missing,
            'addr_long': addr_long,
        }
    # worker 边界：任何异常都落成失败记录返回，绝不能让进程池吞掉
    # noinspection PyBroadException
    except Exception as e:
        import traceback
        return {
            'success': False,
            'filename': filename,
            'error': str(e)[:100],
            'traceback': traceback.format_exc()[-2000:],
        }


# ============================================================================
# 【文件操作代码】压缩包处理
# ============================================================================

def _find_zip_files(folder: str) -> List[str]:
    """扫描文件夹内的 .zip 文件"""
    return [f for f in os.listdir(folder)
            if f.lower().endswith(('.zip',)) and os.path.isfile(os.path.join(folder, f))]


def _safe_extractall(zf: zipfile.ZipFile, dest: str) -> None:
    """带 Zip-Slip 防护的解压：拒绝任何会跳出目标目录的成员路径。
    并兼容 GBK 文件名：未置 UTF-8 标志的 zip（如 Windows 自带压缩）会被 zipfile
    误按 cp437 解码成乱码，这里回退 GBK 重新解码还原中文名。"""
    base = os.path.abspath(dest)
    for info in zf.infolist():
        # 解码后的名字只用于计算落盘路径，绝不能写回 info.filename：
        # Python 3.13 的 zf.open() 会拿该名回查 NameToInfo（键仍是原始名），
        # 改写后会误抛 KeyError（嵌套包 0249-0258 即因此解压失败）。
        name = info.filename
        # 未置 UTF-8 标志(0x800)时按 GBK 重解码；编解码任一失败则保留原名
        if not (info.flag_bits & 0x800):
            try:
                name = name.encode('cp437').decode('gbk')
            except (UnicodeEncodeError, UnicodeDecodeError):
                pass
        # zip 内同时兼容 / 与 \ 分隔符
        rel = name.replace('\\', '/')
        target = os.path.abspath(os.path.join(base, rel))
        if target != base and not target.startswith(base + os.sep):
            raise zipfile.BadZipFile("压缩包含非法路径成员（疑似 Zip-Slip）: " + name)
        # 成员手动落盘：info 保持原样交给 zf.open，避免与库内部索引失配
        if info.is_dir():
            os.makedirs(target, exist_ok=True)
            continue
        parent_dir = os.path.dirname(target)
        if parent_dir:
            os.makedirs(parent_dir, exist_ok=True)
        with zf.open(info) as src_f, open(target, 'wb') as out_f:
            shutil.copyfileobj(src_f, out_f)


def _extract_zip(zip_path: str, extract_to: str) -> bool:
    """解压 zip 到指定目录，返回解压根目录（即压缩包里的顶层文件夹名）"""
    os.makedirs(extract_to, exist_ok=True)
    try:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            _safe_extractall(zf, extract_to)
        return True
    except zipfile.BadZipFile as e:
        print(f"  压缩包损坏: {e}")
        return False
    # 其余解压异常（磁盘/权限/加密等）同样只记录不抛出，保证坏包不阻断整个流程
    # noinspection PyBroadException
    except Exception as e:
        print(f"  解压失败: {e}")
        return False


# 需要忽略的临时/内部目录（自动解压残留、bench 测试等）
_IGNORED_DIRS = frozenset([
    EXTRACTED_FOLDER, "_bench", "_bench2", "_bench_extract",
    "_debug_extract", "_已解压_check", "_temp", "_tmp", "_verify",
    "原始记录汇总", "签名图片", "__pycache__",
])


def _collect_report_batches() -> Tuple[List[Tuple[str, str, str]], str, List[str]]:
    """
    分析检验报告文件夹，返回待处理的批次列表。
    每个批次: (源目录绝对路径, 目标目录绝对路径, 批次显示名)
    
    支持三种输入方式，全部自动识别：
      1. zip 压缩包 → 自动解压后按子文件夹分批（优先）
      2. 手动解压后的子文件夹直接放检验报告里 → 逐个识别
      3. 散放的 docx 文件直接在检验报告根目录（旧逻辑）
    """
    batches = []
    seen_dirs = set()   # 已收集过的绝对路径
    seen_names = set()  # 已收集过的批次名（zip 解压的优先，跳过手动解压的同名）
    bad_zips = []       # 损坏/解压失败的压缩包（汇总时点名）

    def add_batch(src: str, name: str) -> None:
        """统一添加批次，自动生成输出目录，去重"""
        src = os.path.abspath(src)
        if src in seen_dirs:
            return
        # 忽略临时/内部目录（以下划线开头或在忽略列表里）
        base = os.path.basename(src)
        if base.startswith('_') or base in _IGNORED_DIRS:
            return
        # 同名但不同目录：先占坑的批次会让整个目录被跳过，必须显式告警
        # （否则该目录下的报告漏生成且终端无任何提示）
        if name in seen_names:
            print(f"  ！警告：批次名「{name}」与已收集批次重名，目录 {src} 被跳过，"
                  f"请重命名该文件夹后重跑")
            return
        # 必须有报告输入文件（.docx 或待转换的 .doc/.wps）才算有效批次。
        # 递归计数：兼容 zip 内双层嵌套文件夹。
        report_count = len(_list_report_inputs(src, recursive=True))
        if report_count == 0:
            return
        out = os.path.join(OUTPUT_FOLDER, name + "原始记录")
        batches.append((src, out, name))
        seen_dirs.add(src)
        seen_names.add(name)

    # ---- 方式 1: 处理 zip（优先级最高） ----
    zip_files = _find_zip_files(REPORTS_FOLDER)
    extract_root = os.path.join(REPORTS_FOLDER, EXTRACTED_FOLDER)

    # 非 zip 压缩包（.7z/.rar）程序不会自动解包，必须显式告警提醒人工核对份数
    other_archives = [f for f in os.listdir(REPORTS_FOLDER)
                      if f.lower().endswith(('.7z', '.rar'))
                      and os.path.isfile(os.path.join(REPORTS_FOLDER, f))]
    if other_archives:
        print(f"  ！发现 {len(other_archives)} 个非 zip 压缩包（.7z/.rar），程序不会自动解包，"
              f"请人工解压或核对份数:")
        for a in sorted(other_archives):
            print(f"      {a}")

    if zip_files:
        print(f"检测到 {len(zip_files)} 个压缩包，准备自动解压...")
        # 清理上次残留
        if os.path.exists(extract_root):
            shutil.rmtree(extract_root, ignore_errors=True)
        os.makedirs(extract_root, exist_ok=True)

        for zname in zip_files:
            zpath = os.path.join(REPORTS_FOLDER, zname)
            # 每个 zip 单独解压到 _已解压/<zip名>/，避免多个 zip 散文件混放，
            # 也保证散文件批次拿到合法批次名（zip 名），不会被当“_已解压”跳过
            zstem = os.path.splitext(zname)[0]
            target = os.path.join(extract_root, zstem)
            os.makedirs(target, exist_ok=True)
            print(f"  解压: {zname}")
            if _extract_zip(zpath, target):
                print(f"  解压完成: {zname}")
            else:
                print(f"  跳过损坏的压缩包: {zname}")
                bad_zips.append(zname)
                continue

        # 注意：嵌套压缩包（zip 内再含 zip/.7z/.rar）不自动递归解压，
        # 解包后统一点名告警、直接略过（见下方 stray 扫描）。

        # 每个 zip 目录内：有子文件夹则按子文件夹分批；直接散文件则用 zip 名做批次名
        #（这些批次名先占坑，后面手动解压同名的会被跳过）
        for zentry in sorted(os.listdir(extract_root)):
            zfull = os.path.join(extract_root, zentry)
            if not os.path.isdir(zfull):
                continue
            inner_dirs = [e for e in sorted(os.listdir(zfull))
                          if os.path.isdir(os.path.join(zfull, e))]
            for d in inner_dirs:
                add_batch(os.path.join(zfull, d), d)
            if not inner_dirs and _list_report_inputs(zfull, recursive=False):
                add_batch(zfull, zentry)

        # 嵌套压缩包（顶层 zip 解出的内容里再出现的 zip/.7z/.rar）一律不自动解压、
        # 直接略过，只点名位置提醒人工核对份数；extract_root 内的档案必然来自顶层 zip 解包
        stray = []
        for root, _dirs, files in os.walk(extract_root):
            for f in files:
                if f.lower().endswith(('.zip', '.7z', '.rar')):
                    fp = os.path.abspath(os.path.join(root, f))
                    stray.append(os.path.relpath(fp, extract_root))
        if stray:
            print(f"  ！发现 {len(stray)} 个嵌套压缩包（程序不会自动解压，已全部略过；"
                  f"如需其中报告请人工解压后重跑，并请核对份数）:")
            for a in sorted(stray):
                print(f"      {a}")

    # ---- 方式 2: 检验报告里已有的子文件夹（手动解压的） ----
    # 如果 zip 解压已经占了同名的，这里会被自动跳过
    for entry in sorted(os.listdir(REPORTS_FOLDER)):
        full = os.path.join(REPORTS_FOLDER, entry)
        if os.path.isdir(full):
            add_batch(full, entry)

    # ---- 方式 3: 检验报告根目录直接散放报告文件（旧逻辑） ----
    root_inputs = _list_report_inputs(REPORTS_FOLDER, recursive=False)
    if root_inputs:
        # 根批次绝不能递归，否则会把子文件夹/已解压目录整棵重复生成一轮
        # 只认直接层文件，批次名用"检验报告"
        src_root = os.path.abspath(REPORTS_FOLDER)
        if src_root not in seen_dirs:
            report_count = len(root_inputs)
            if report_count:
                out = os.path.join(OUTPUT_FOLDER, REPORTS_FOLDER + "原始记录")
                batches.append((src_root, out, REPORTS_FOLDER))
                seen_dirs.add(src_root)
                seen_names.add(REPORTS_FOLDER)

    return batches, extract_root, bad_zips


def _cleanup_extracted(extract_root: str) -> None:
    """清理临时解压目录"""
    if os.path.exists(extract_root):
        try:
            shutil.rmtree(extract_root, ignore_errors=True)
            print(f"已清理临时解压目录")
        except OSError as e:
            print(f"清理临时目录失败（不影响结果）: {e}")


def _clean_output_folder() -> List[Tuple[str, OSError]]:
    """每次运行生成前清空成品目录，防止上一批旧批次目录残留与本次成品混杂。

    只删 OUTPUT_FOLDER 内的子项、不删目录本身（与 process_reports 开头
    “确保成品根目录存在”的职责不重叠）；.gitkeep 占位文件保留。
    返回删除失败的 (名称, 异常) 列表——Windows 上成品被 WPS/Word 打开时
    会占用删不掉，调用方必须点名并中止运行：新旧混杂比不生成更危险。
    """
    failed: List[Tuple[str, OSError]] = []
    if not os.path.isdir(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER, exist_ok=True)
        return failed
    for name in os.listdir(OUTPUT_FOLDER):
        if name == '.gitkeep':
            continue
        path = os.path.join(OUTPUT_FOLDER, name)
        try:
            if os.path.isdir(path) and not os.path.islink(path):
                shutil.rmtree(path)
            else:
                os.remove(path)
        except OSError as e:
            failed.append((name, e))
    return failed


def _list_report_docx(folder: str, recursive: bool = False) -> List[str]:
    """列出目录下所有有效的检验报告 .docx 文件。

    recursive=True 时递归子目录，返回相对 folder 的路径（sorted）。
    """
    if recursive:
        out = []
        for root, _dirs, files in os.walk(folder):
            for f in files:
                if _is_valid_report_docx(f):
                    out.append(os.path.relpath(os.path.join(root, f), folder))
        return sorted(out)
    try:
        all_files = os.listdir(folder)
    except OSError:
        return []
    return sorted([f for f in all_files if _is_valid_report_docx(f)])


def _preprocess_reports(folder: Optional[str] = None, recursive: bool = False) -> bool:
    """扫描文件夹，把非 .docx 的报告(.doc/.wps/.docm)批量转成 .docx。

    recursive=True 时递归处理全部子目录（用于 zip 内嵌套结构）。
    """
    target = folder if folder else REPORTS_FOLDER
    valid_exts = ('.doc', '.wps', '.docm')

    candidates = []
    if recursive:
        for root, _dirs, files in os.walk(target):
            for f in files:
                if f.startswith('~$') or '合格证' in f:
                    continue
                if f.lower().endswith(valid_exts):
                    src = os.path.abspath(os.path.join(root, f))
                    candidates.append((src, os.path.splitext(src)[0] + ".docx"))
    else:
        for f in os.listdir(target):
            if f.startswith('~$') or '合格证' in f:
                continue
            if f.lower().endswith(valid_exts):
                src = os.path.abspath(os.path.join(target, f))
                candidates.append((src, os.path.splitext(src)[0] + ".docx"))

    if not candidates:
        return True

    print(f"检测到 {len(candidates)} 份非 .docx 检验报告，正在转换...")
    word = None
    try:
        import win32com.client
        word = win32com.client.Dispatch("KWps.Application")
        word.Visible = False
        try:
            word.DisplayAlerts = 0  # wdAlertsNone：损坏/加密文件不弹模态框卡死整批
        # COM 属性赋值异常不可控，仅表示弹窗抑制未生效
        # noinspection PyBroadException
        except Exception:
            pass
        converted = 0
        for src, dst in candidates:
            # 已存在同名 docx 就跳过，避免覆盖用户新手动改的
            if os.path.exists(dst):
                print(f"  跳过（已存在）: {os.path.basename(dst)}")
                continue
            try:
                doc = word.Documents.Open(src)
                doc.SaveAs2(dst, FileFormat=16)  # 16 = wdFormatXMLDocument (.docx)
                doc.Close()
                converted += 1
                print(f"  已转换: {os.path.basename(src)} -> {os.path.basename(dst)}")
            # COM 转换单文件失败（损坏/加密等）只记录，继续下一份
            # noinspection PyBroadException
            except Exception as e:
                print(f"  转换失败 {os.path.basename(src)}: {e}")
        word.Quit()
        print(f"转换完成：成功 {converted}/{len(candidates)}")
        return True
    except ImportError:
        print("未安装 pywin32 或 WPS，跳过自动转换")
        return False
    # COM 环境级失败必须接住并优雅返回，保证程序可继续
    # noinspection PyBroadException
    except Exception as e:
        print(f"转换过程出错: {e}")
        if word is not None:
            try:
                word.Quit()
            # noinspection PyBroadException
            except Exception:
                pass
        return False


def _run_task(args: Tuple[str, str, str, str, Optional[int]]) -> Dict[str, Any]:
    """多进程池的 worker：只负责生成单份原始记录。"""
    filename, src_folder, output_folder, batch_name, appendix_page = args
    result = process_single_report((filename, src_folder, output_folder, appendix_page))
    result['batch_name'] = batch_name
    result['src_folder'] = src_folder
    return result


_APPENDIX_HEADING_FULL = "道路运输液体危险货物罐式车辆充装介质附页"


def _appendix_page_in_section(word, report_path: str) -> Optional[int]:
    """用 WPS 排版引擎求附页标题在其所在节中的节内页码（1 起）。

    节内页码 = 标题物理页 - 该节起始物理页 + 1。用于成品 pgNumType@start，
    使附页页码与源报告完全一致。
    """
    doc = word.Documents.Open(os.path.abspath(report_path), ReadOnly=True)
    try:
        doc.Repaginate()
        sec_start = doc.Sections(doc.Sections.Count).Range
        sec_start.Collapse(1)  # wdCollapseStart
        page_section = sec_start.Information(3)  # wdActiveEndPageNumber

        # 标题文字也可能出现在前面的目录页；遍历所有匹配，取落在最后一节内
        # （物理页 >= 本节起始页）且页码最大的一处，即真正的附页标题。
        sel = word.Selection
        sel.HomeKey(Unit=6)  # wdStory
        page_heading = None
        finder = sel.Find
        finder.ClearFormatting()
        finder.Text = _APPENDIX_HEADING_FULL
        try:
            finder.Wrap = 0  # wdFindStop：到文档末尾即停，禁止绕回开头造成死循环
        # COM 属性可能不支持，默认行为也安全
        # noinspection PyBroadException
        except Exception:
            pass
        while finder.Execute():
            p = sel.Information(3)
            if p >= page_section and (page_heading is None or p > page_heading):
                page_heading = p
            sel.Collapse(0)  # wdCollapseEnd，从匹配末尾继续向后找
        if page_heading is None:
            return None
        # int() 包裹：COM Information 返回 Any，显式转 int 后再做减法
        return int(page_heading) - int(page_section) + 1
    finally:
        doc.Close(False)


def _acquire_single_instance() -> Optional[int]:
    """命名互斥量：防止两个 main 同时运行（会互删 _已解压、互踩成品）。

    返回互斥量句柄（进程退出时由 OS 自动释放）；已有实例在运行则返回 None。
    """
    if os.name != 'nt':
        return -1
    try:
        import ctypes
        k32 = ctypes.WinDLL('kernel32', use_last_error=True)
        error_already_exists = 183
        handle = k32.CreateMutexW(None, True, "Global\\yuanshijilu_main_singleton_v1")
        if not handle:
            return -1  # 创建失败不阻塞业务，仅失去单实例保护
        if ctypes.get_last_error() == error_already_exists:
            return None
        return handle
    except (OSError, AttributeError, TypeError, ValueError):
        return -1


def process_reports() -> int:
    # 【文件操作代码】检查文件夹是否存在
    if not os.path.exists(REPORTS_FOLDER):
        print(f"文件夹 '{REPORTS_FOLDER}' 不存在")
        return 1

    # 【文件操作代码】创建输出文件夹
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)

    # ---- Tee 一开始就开，保证从第一行 print 开始全部进日志 ----
    os.makedirs(LOG_FOLDER, exist_ok=True)
    _tmp_path = os.path.join(LOG_FOLDER, f"_pending_{os.getpid()}.log")
    _log = open(_tmp_path, 'w', encoding='utf-8')
    sys.stdout = _Tee(sys.__stdout__, _log)
    _final_box = [_tmp_path]   # 用列表包一层，内层函数可修改
    problems = 0
    try:
        problems = _process_body(_log, _tmp_path, _final_box)
    finally:
        sys.stdout = sys.__stdout__
        try:
            _log.close()
        except OSError:
            pass
        # 确保最终文件名存在（如果 Windows 上 rename 失败，tmp_path 已经是最终内容）
        final = _final_box[0]
        if final != _tmp_path and os.path.exists(_tmp_path) and not os.path.exists(final):
            try:
                os.replace(_tmp_path, final)
            except OSError:
                pass
        print(f"\n📄 本次运行日志已保存: {final}")
    return problems


def _process_body(log, tmp_path: str, final_box: List[str]) -> int:
    """process_reports 的实际处理主体。返回问题数量（0 表示全部正常）。"""
    start_time = time.time()

    # 1. 收集批次（解压 zip 或直接用检验报告根目录）
    batches, extract_root, bad_zips = _collect_report_batches()

    if not batches:
        print("没有可处理的检验报告文件")
        _cleanup_extracted(extract_root)
        if bad_zips:
            print(f"另有 {len(bad_zips)} 个损坏压缩包未处理：{', '.join(bad_zips)}")
        return len(bad_zips)

    # 2. 先对每个批次做预处理 + 收集任务（这步单线程，很快）
    all_tasks = []                  # [(filename, src_folder, output_folder, batch_name), ...]
    total_count = 0
    all_convert_failed = []         # [(批次名, 源文件相对路径), ...] 转换失败导致必缺份

    print(f"\n共识别出 {len(batches)} 个批次，正在收集任务...")

    # 生成前清空上一批成品：批次已确认有效（batches 非空），旧批次目录再保留只会
    # 与本次成品混杂（如上次放 6-8 月 zip、本次只放 3-5 月）。有成品被占用删不掉时
    # 中止本次运行，避免新旧混杂交付
    locked = _clean_output_folder()
    if locked:
        print(f"\n！清空“{OUTPUT_FOLDER}”失败：{len(locked)} 个旧项目被占用"
              f"（很可能正在 WPS/Word 中打开），请全部关闭后重跑：")
        for name, err in locked:
            print(f"      {name}（{err}）")
        _cleanup_extracted(extract_root)
        return 1
    print(f"已清空“{OUTPUT_FOLDER}”中的上次成品（.gitkeep 保留）")

    for src, out, name in batches:
        # 根目录散文件批次只收直接层文件：递归会把 zip 解压内容和子文件夹整树重复生成
        is_root_batch = os.path.abspath(src) == os.path.abspath(REPORTS_FOLDER)
        rec = not is_root_batch
        # 转换前留底：.doc/.wps 输入清单，供转换后核对漏转
        legacy_inputs = [f for f in _list_report_inputs(src, recursive=rec)
                         if f.lower().endswith(('.doc', '.wps', '.docm'))]

        _preprocess_reports(src, recursive=rec)

        report_files = _list_report_docx(src, recursive=rec)

        # fail-closed：转换后拿不到对应 .docx 的源文件逐个点名（否则会静默缺份）。
        # 必须先算漏转再判空跳过——整批 .wps 转换失败时 report_files 为空，
        # 若先 continue 会把整批丢失藏成一句"没有找到检验报告，跳过"且退出码仍为 0。
        docx_stems = {os.path.splitext(f)[0] for f in report_files}
        convert_failed = sorted(
            f for f in legacy_inputs
            if os.path.splitext(f)[0] not in docx_stems)
        if convert_failed:
            print(f"  ！批次「{name}」有 {len(convert_failed)} 份源文件未能转为 .docx（对应成品将缺失）:")
            for f in convert_failed:
                print(f"      {f}")
            all_convert_failed.extend((name, f) for f in convert_failed)

        if not report_files:
            print(f"  批次「{name}」目录下没有找到检验报告，跳过")
            continue

        os.makedirs(out, exist_ok=True)

        total_count += len(report_files)

        for filename in report_files:
            all_tasks.append((filename, src, out, name, None))

        print(f"  批次「{name}」→ {len(report_files)} 份 → {os.path.basename(out)}")

    if not all_tasks:
        print("没有可处理的检验报告文件")
        _cleanup_extracted(extract_root)
        return len(bad_zips) + len(all_convert_failed)

    # ---- all_tasks 齐了 → 算出正确的日志文件名，安全地重命名 ----
    all_filenames = [t[0] for t in all_tasks]
    final_name = _build_log_filename(all_filenames)
    final_path = os.path.join(LOG_FOLDER, final_name)
    final_box[0] = final_path

    # Windows 上文件开着不能 rename → 先关、rename、再 reopen
    sys.stdout = sys.__stdout__
    log.flush()
    log.close()
    try:
        os.replace(tmp_path, final_path)
    except OSError:
        # rename 失败（比如已存在同名文件），就用 append 打开目标文件继续写
        pass
    log = open(final_path, 'a', encoding='utf-8')
    sys.stdout = _Tee(sys.__stdout__, log)

    # 3. 全部任务合并进一个多进程池，进程池只创建一次！
    print(f"\n共 {total_count} 份检验报告，使用 8 进程并行处理...")
    process_count = 8

    success_count = 0
    all_special_media = []
    all_guangxi_plate = []
    all_appendix = []
    appendix_missing = []
    failed_files = []
    addr_long_files = []
    generated = set()       # (批次名, 源文件名) 首轮确实写出成品的
    expected = {(t[3], t[0]) for t in all_tasks}

    print(f"  (已提交 {total_count} 份，开始流式生成...)\n")
    appendix_results = []   # 首轮生成成功且含附页的结果，待补页码后二次生成
    try:
        with multiprocessing.Pool(processes=process_count) as pool:
            for result in pool.imap_unordered(_run_task, all_tasks):
                batch = str(result.get('batch_name') or '')
                if os.path.isabs(batch):
                    batch = os.path.basename(batch.rstrip('\\/'))
                if result['success']:
                    success_count += 1
                    generated.add((batch, result['filename']))
                    rel = result['rel_output']
                    if result['special_media']:
                        all_special_media.append(rel)
                    if result['guangxi_plate']:
                        all_guangxi_plate.append(rel)
                    if result.get('addr_long'):
                        addr_long_files.append(rel)
                    src_rel = os.path.join(batch, result['filename'])
                    if result.get('has_appendix'):
                        all_appendix.append((src_rel, rel))
                        appendix_results.append(result)
                    if result.get('appendix_missing'):
                        appendix_missing.append((src_rel, rel))
                    print(f"  [完成✓] {result['filename']}")
                else:
                    print(f"  [失败] {result['filename']} - {result['error']}")
                    tb = result.get('traceback')
                    if tb:
                        for line in str(tb).splitlines():
                            print("    " + line)
                    failed_files.append(os.path.join(batch, result['filename']))

        # ---- 附页第二轮：用 WPS 排版求源附页节内页码，带页码重新生成 ----
        appendix_regen_failed = []
        appendix_page_unknown = []   # 页码求不到、保持自动编号的成品（需 verify.py 复核）
        if appendix_results:
            print(f"\n共 {len(appendix_results)} 份含充装介质附页，"
                  f"正在用 WPS 核对源报告页码后重新生成（保证页眉/页码与源页一致）...")
            import win32com.client
            word = None
            paged_tasks = []
            try:
                word = win32com.client.Dispatch("KWps.Application")
                word.Visible = False
                try:
                    word.DisplayAlerts = 0
                # COM 属性异常不影响后续转换
                # noinspection PyBroadException
                except Exception:
                    pass
                for r in appendix_results:
                    src_path = os.path.join(r['src_folder'], r['filename'])
                    b = str(r.get('batch_name') or '')
                    if os.path.isabs(b):
                        b = os.path.basename(b.rstrip('\\/'))
                    try:
                        pn = _appendix_page_in_section(word, src_path)
                    # COM 求页码失败只让该份保持自动编号，不能中断整批
                    # noinspection PyBroadException
                    except Exception as e:
                        print(f"  [页码获取失败] {r['filename']} - {e}")
                        pn = None
                    if pn is None:
                        appendix_page_unknown.append(os.path.join(b, r['filename']))
                    out_folder = os.path.dirname(r['output_path'])
                    paged_tasks.append((r['filename'], r['src_folder'],
                                        out_folder, r['batch_name'], pn))
                    print(f"  附页页码 {pn if pn else '?'} ← {r['filename']}")
            # WPS 整体不可用时降级为自动页码，附页内容不受影响；
            # 但这批成品的页码全部未核对，必须逐份点名交 verify.py 复核
            # noinspection PyBroadException
            except Exception as e:
                print(f"  WPS 不可用，附页页码保持自动编号（不影响内容）: {e}")
                appendix_page_unknown.extend(
                    os.path.join(str(r.get('batch_name') or ''), r['filename'])
                    for r in appendix_results)
            finally:
                if word is not None:
                    try:
                        word.Quit()
                    # noinspection PyBroadException
                    except Exception:
                        pass

            if paged_tasks:
                with multiprocessing.Pool(processes=process_count) as pool:
                    for result in pool.imap_unordered(_run_task, paged_tasks):
                        if not result['success']:
                            b = str(result.get('batch_name') or '')
                            if os.path.isabs(b):
                                b = os.path.basename(b.rstrip('\\/'))
                            print(f"  [附页二次生成失败] {result['filename']} - {result['error']}")
                            appendix_regen_failed.append(os.path.join(b, result['filename']))
                print("附页页码核对完成")
    finally:
        _cleanup_extracted(extract_root)

    # 全局汇总
    print(f"\n{'='*50}")
    print(f"全部完成！共生成 {success_count} 份原始记录")
    print(f"{'='*50}")

    # —— 交付前最重要的名单：哪些源报告没有生成出成品（用户据此与对方核对份数）——
    not_generated = sorted(os.path.join(b, f) for b, f in (expected - generated))
    if not_generated:
        print(f"\n★未生成成品的源报告（共 {len(not_generated)} 份，发成品前必须处理/补发）：")
        for f in not_generated:
            print(f"  {f}")
    else:
        print("\n★源报告 %d 份全部生成出成品，无遗漏。" % len(expected))

    if all_convert_failed:
        print(f"\n源文件未能转为 .docx（共 {len(all_convert_failed)} 份，成品必然缺失）：")
        for b, f in all_convert_failed:
            print(f"  {b}\\{f}")

    if bad_zips:
        print(f"\n损坏/解压失败的压缩包（共 {len(bad_zips)} 个，其中报告全部未处理）：")
        for z in bad_zips:
            print(f"  {z}")

    if appendix_regen_failed:
        print(f"\n附页二次生成失败（成品内容已在，但页码可能与源页不一致，"
              f"共 {len(appendix_regen_failed)} 份，请重跑或人工核对）：")
        for f in appendix_regen_failed:
            print(f"  {f}")

    if appendix_page_unknown:
        print(f"\n附页页码未能核对（共 {len(appendix_page_unknown)} 份，成品附页页码为自动编号，"
              f"可能与源页不一致，请运行 verify.py 复核）：")
        for f in appendix_page_unknown:
            print(f"  {f}")

    if all_special_media:
        print(f"\n装运介质不是汽油/柴油的有（共 {len(all_special_media)} 份，路径相对于“{OUTPUT_FOLDER}”）：")
        for f in sorted(all_special_media):
            print(f"  {f}")

    if all_guangxi_plate:
        print(f"\n机动车号牌是桂开头的有（共 {len(all_guangxi_plate)} 份，路径相对于“{OUTPUT_FOLDER}”）：")
        for f in sorted(all_guangxi_plate):
            print(f"  {f}")

    if addr_long_files:
        print(f"\n★单位地址超长（共 {len(addr_long_files)} 份，模板行高固定，PDF 可能截断，请手动调整行高）：")
        print(f"  格式：成品位置（相对于“{OUTPUT_FOLDER}”）")
        for f in sorted(addr_long_files):
            print(f"  {f}")

    if all_appendix:
        print(f"\n含充装介质附页的报告（附页已原样追加到原始记录末尾，共 {len(all_appendix)} 份）：")
        print(f"  格式：源报告位置（“检验报告”内批次\\文件名）  →  成品位置（相对于“{OUTPUT_FOLDER}”）")
        for src, out in sorted(all_appendix):
            print(f"  {src}  →  {out}")

    if appendix_missing:
        print(f"\n装运介质写明“见附页”但报告中未找到附页正文（共 {len(appendix_missing)} 份，请人工核查）：")
        print(f"  格式：源报告位置（“检验报告”内批次\\文件名）  →  成品位置（相对于“{OUTPUT_FOLDER}”）")
        for src, out in sorted(appendix_missing):
            print(f"  {src}  →  {out}")

    if failed_files:
        print(f"\n处理失败的有（共 {len(failed_files)} 份，路径为“检验报告”内批次文件夹\\源报告文件名）：")
        for f in sorted(failed_files):
            print(f"  {f}")

    end_time = time.time()
    print(f"\n总用时 {end_time - start_time:.2f} 秒，平均 {((end_time - start_time)/success_count if success_count else 0):.2f} 秒/份")

    return (len(not_generated) + len(all_convert_failed)
            + len(bad_zips) + len(appendix_regen_failed)
            + len(appendix_page_unknown))


if __name__ == "__main__":
    _mutex = _acquire_single_instance()
    if _mutex is None:
        print("已有一个 main.py 正在运行，为避免互删解压目录/互踩成品，本次启动已退出。")
        sys.exit(2)
    _code = process_reports()
    sys.exit(1 if _code else 0)
