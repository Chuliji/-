# -*- coding: utf-8 -*-
"""自动注册本地 _deps 依赖（pywin32 等），用于 .wps/.doc 报告自动转换"""
import os as _os, sys as _sys


def _get_deps_dir():
    """获取 _deps 目录路径，兼容开发模式和 PyInstaller 打包。"""
    # 1. PyInstaller onefile 解压目录
    if hasattr(_sys, "_MEIPASS"):
        _p = _os.path.join(_sys._MEIPASS, "_deps")
        if _os.path.isdir(_p):
            return _p
    # 2. exe 同级目录（onedir 模式）
    _exe_dir = _os.path.dirname(_os.path.abspath(_sys.executable))
    _p = _os.path.join(_exe_dir, "_deps")
    if _os.path.isdir(_p):
        return _p
    # 3. 脚本同级目录（开发模式）
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

from typing import Any, Dict, Iterable, List, Match, Optional, Set, Tuple

from docx import Document
from docx.oxml.ns import qn
import lxml.etree as etree
import os
import re
import sys
import time
import zipfile
import shutil
import multiprocessing
import html
import math
import random

# 文本索引表：(扁平单元格文本列表, 列数, 行数)
GridTbl = Tuple[List[str], int, int]

LOG_FOLDER = os.path.join("终端结果", "check运行日志")


_ANSI_RE = re.compile(r'\x1b\[[0-9;]*m')
_RED = '\x1b[91m'
_RESET = '\x1b[0m'


def _build_log_filename(all_filenames):
    """与 main.py 完全一致的命名：按报告编号范围取名（前面再加 check_）。

    同一批原始记录重复检查时日志名相同、直接覆盖，不再产生一堆时间戳日志。
    """
    prefixes = set()
    suffix_nums = []
    for fn in all_filenames:
        m = re.search(r'(B-[A-Z]+)(\d+)', fn)
        if m:
            prefixes.add(m.group(1))
            suffix_nums.append(int(m.group(2)))

    if not suffix_nums:
        return "check_run_" + time.strftime('%Y%m%d_%H%M%S') + ".log"

    prefix = sorted(prefixes)[0] if len(prefixes) == 1 else "B-ALL"
    lo, hi = min(suffix_nums), max(suffix_nums)
    lo_str, hi_str = str(lo), str(hi)
    if lo == hi:
        return f"check_{prefix}{lo_str}.log"
    common_len = 0
    for a, b in zip(lo_str, hi_str):
        if a == b:
            common_len += 1
        else:
            break
    return f"check_{prefix}{lo_str}-{hi_str[common_len:]}.log"

# 逐字段比对清单（与 main.py 提取字段对齐）
FIELDS = ["报告编号", "使用单位", "机动车号牌", "单位地址", "道路运输证号", "总质量", "核定载质量",
          "制造企业", "制造日期", "设计代码", "产品标准", "产品型号",
          "VIN码", "罐体编号", "罐体容积", "罐体外形尺寸", "分仓数量",
          "封头材质", "筒体材质", "封头厚度", "筒体厚度", "装运介质",
          "封头1", "封头2", "封头3", "封头4", "封头17", "封头18", "封头19", "封头20",
          "筒体5", "筒体6", "筒体7", "筒体8", "筒体9", "筒体10",
          "筒体11", "筒体12", "筒体13", "筒体14", "筒体15", "筒体16",
          "检验日期", "审核日期", "下次检验日期"]

# 检验报告必填字段：报告侧若漏填（空值），原始记录必然同步空白，
# 普通的“两边比对”会因两边都空而放行，因此必须单独判空。
# 适装介质列表不列入：多数报告整表为“——”占位属正常情况；
# 报告若实际填了介质，仍由下方两边比对（FIELDS 之外的列表比对）保证不丢不错。
# 如业务上个别字段允许留空（例如暂时未审核），直接从本清单删除对应项即可。
REQUIRED_FIELDS = list(FIELDS)

# 这些字段提取时只 strip 首尾、保留中间原文（未做去空格处理），
# 可直接检测报告录入时误入的连续空格/全角空格（脏数据会被原样拷贝，比对查不出）。
_SPACE_CHECK_FIELDS = ["使用单位", "单位地址", "制造企业",
                       "封头材质", "筒体材质", "装运介质"]
_SPACE_RE = re.compile(r'\u3000|[ \t]{2,}')


class _Tee:
    """同时写终端和日志文件的 stdout 代理。

    终端流保留 ANSI 颜色码；写日志文件时自动剥离颜色码，保证日志纯文本。
    """
    def __init__(self, *streams):
        self.streams = streams

    def write(self, data):
        for s in self.streams:
            payload = data if s is sys.__stdout__ else _ANSI_RE.sub('', data)
            s.write(payload)
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


def _norm(s):
    """去掉字符串内全部空白（含全角空格），用于号牌/编号/日期/数值等字段。"""
    return re.sub(r'[\s\u3000]+', '', s or '')


def _build_doc_index(doc) -> Tuple[List[GridTbl], List[str]]:
    """单遍物化整份文档：一次性读入所有表格的单元格文本和段落文本。

    返回 (tables, paras)：
      tables = [(grid, ncol, nrow), ...]，grid 为扁平的原始单元格文本列表
      paras  = [段落文本, ...]
    后续所有提取都只读这份内存索引，避免对同一批单元格反复调用昂贵的
    Cell.text（每次都会重新拼接段落/run 文本）。
    """
    tables = []
    # python-docx 未公开网格 API，仅此处读取一次内部成员，后续不再触碰
    # noinspection PyProtectedMember
    for table in doc.tables:
        ncol = table._column_count
        grid = [cell.text for cell in table._cells]
        tables.append((grid, ncol, len(table.rows)))
    paras = [p.text for p in doc.paragraphs]
    return tables, paras


def _gt(tbl: GridTbl, row: int, col: int) -> str:
    """从文本索引表 (grid, ncol, nrow) 里取去首尾空白后的单元格文本。"""
    grid, ncol, _ = tbl
    try:
        return grid[row * ncol + col].strip()
    except (IndexError, TypeError):
        return ""


def _find_tbl(tables: List[GridTbl], keywords: Iterable[str]) -> Optional[GridTbl]:
    """按内容定位表格：返回第一张包含任一关键词的文本索引表（不依赖表格序号）。"""
    for tbl in tables:
        for t in tbl[0]:
            for kw in keywords:
                if kw in t:
                    return tbl
    return None


_DATE_RE = re.compile(r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日')


def _fmt_date(m: Match[str]) -> str:
    return "%d年%02d月%02d日" % (int(m.group(1)), int(m.group(2)), int(m.group(3)))


def _extract_dates(tables: List[GridTbl]) -> Tuple[str, str, str]:
    """单遍扫描所有表格，同时提取检验/审核签名日期与下次检验日期。

    返回 (检验日期, 审核日期, 下次检验日期)，兼容日期与标签同格/分格两种版式。
    """
    jy = sh = nxt = ""
    for grid, ncol, nrow in tables:
        for r in range(nrow):
            comp = [_norm(grid[r * ncol + c]) for c in range(ncol)]
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
                if not nxt and "下次检验日期" in t:
                    m = re.search(r'下次检验日期[：:]?\s*(\d{4}\s*年\s*\d{1,2}\s*月)', t)
                    if m:
                        nxt = _norm(m.group(1))
    return jy, sh, nxt


def _extract_medium_list(tbl: GridTbl) -> List[Dict[str, str]]:
    """按“表头列名”提取适装介质列表（兼容 9/10 列两种版式）。入参为文本索引表。"""
    grid, ncol, nrow = tbl
    hdr_r = None
    for r in range(nrow):
        row_comp = [_norm(grid[r * ncol + c]) for c in range(ncol)]
        if any(x == "序号" for x in row_comp) and any("介质名称" in x for x in row_comp):
            hdr_r = r
            break
    if hdr_r is None:
        return []
    hdr = [_norm(grid[hdr_r * ncol + c]) for c in range(ncol)]

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
        return "" if c < 0 else _gt(tbl, r, c)

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


def _extract_thickness(tbl: GridTbl) -> Dict[str, str]:
    """按“点数→数值”解析测厚表（封头/筒体都可，兼容 18/19 列）。入参为文本索引表。"""
    grid, ncol, nrow = tbl
    data = {}
    int_re = re.compile(r'^\d+$')
    dec_re = re.compile(r'^\d+\.\d+$')
    for r in range(nrow):
        section = None
        pending = None
        for c in range(ncol):
            t = _norm(grid[r * ncol + c])
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


def _el_text(el):
    return "".join(t.text or "" for t in el.iter(qn('w:t')))


def _apd_boundary(kids) -> Optional[Tuple[int, int]]:
    """定位附页节边界 (start, end)；无附页返回 None。

    主路径：含“充装以下介质”标记的整表，起点回溯“…充装介质附页”标题段，
    保证标题/报告编号/表格/注释完整；终点止于 body 末尾 sectPr 之前。
    回退：旧版成品中附页为正文段落（分页符之后）。
    """
    end = len(kids) - 1 if kids and kids[-1].tag == qn('w:sectPr') else len(kids)

    ti = None
    for i, kid in enumerate(kids):
        if kid.tag == qn('w:tbl') and "充装以下介质" in _el_text(kid):
            ti = i  # 与 main.py 一致取最后一个：真正附页在文末，正文前表可能含同字样
    if ti is not None:
        start = ti
        for i in range(ti - 1, -1, -1):
            kid = kids[i]
            if kid.tag != qn('w:p'):
                break
            if "充装介质附页" in _el_text(kid):
                start = i
                break
        return start, end

    # 回退：旧版成品中附页是正文段落（分页符之后），同样取最后一处
    pi = None
    for i, kid in enumerate(kids):
        if kid.tag == qn('w:p') and "充装以下介质" in _el_text(kid):
            pi = i
    if pi is not None:
        return pi, end
    return None


# noinspection PyProtectedMember
def _extract_appendix_text(doc) -> str:
    """提取充装介质附页整节文本（去全部空白），无附页返回 ''。"""
    # 快速预检：正文不含标记时直接返回（大多数文件的常态路径）
    full = "".join(t.text or "" for t in doc.element.body.iter(qn('w:t')))
    if "充装以下介质" not in full and "充装介质附页" not in full:
        return ""
    kids = list(doc.element.body)
    b = _apd_boundary(kids)
    if b is None:
        return ""
    start, end = b
    return _norm("".join(_el_text(kids[j]) for j in range(start, end)))


# noinspection PyProtectedMember
def _appendix_image_hashes(doc) -> Set[str]:
    """收集附页节内正文图片（印章等）的 SHA-256 哈希集合。

    与 main 的整节拷贝对称：只统计附页边界内 r:embed/r:id/r:link 引用的部件。
    """
    import hashlib
    kids = list(doc.element.body)
    b = _apd_boundary(kids)
    if b is None:
        return set()
    start, end = b
    hashes = set()
    for j in range(start, end):
        for node in kids[j].iter():
            for attr in ("embed", "id", "link"):
                rid = node.get(qn("r:" + attr))
                if not rid or rid not in doc.part.rels:
                    continue
                rel = doc.part.rels[rid]
                if rel.is_external:
                    continue
                hashes.add(hashlib.sha256(rel.target_part.blob).hexdigest())
    return hashes


# noinspection PyProtectedMember
def _last_section_effective_refs(doc) -> Dict[Tuple[str, str], Optional[str]]:
    """推导文档最后一节“实际生效”的页眉/页脚引用 {(kind, wtype): rid}（含继承）。

    算法与 main.py 的 _effective_refs 保持一致：沿段落 sectPr 链累积，
    body 末尾 sectPr 最后覆盖。附页在文档最后一节，其页眉/页脚即成品核对对象。
    """
    sect_els = [p.find(qn('w:pPr') + '/' + qn('w:sectPr'))
                for p in doc.element.body.iter(qn('w:p'))]
    sect_els = [s for s in sect_els if s is not None]
    sect_els.append(doc.element.body.find(qn('w:sectPr')))
    refs = {}
    for sect in sect_els:
        if sect is None:
            continue
        for tag in ('w:headerReference', 'w:footerReference'):
            for ref in sect.findall(qn(tag)):
                kind = 'header' if tag == 'w:headerReference' else 'footer'
                refs[(kind, ref.get(qn('w:type')) or 'default')] = ref.get(qn('r:id'))
    return refs


# 页码域识别：仅匹配域指令的第一个域名为 PAGE（不能误匹配 NUMPAGES）
_PG_FIELD_NAME_RE = re.compile(r'^\s*PAGE\b', re.IGNORECASE)


def _strip_page_number_fields(xml_root) -> int:
    """就地删除页眉/页脚 XML 中全部 PAGE 页码域，返回删除元素数。

    main.py 中有同名镜像实现（成品侧剔除），本处供核对侧剔除后再取文本，
    保证“源有页码、成品无页码”是设计行为而非文字错漏。修改口径时两边同步。
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
def _appendix_chrome(doc) -> Tuple[Set[str], Set[str]]:
    """附页末节“真正渲染生效”的页眉/页脚核对数据。

    返回 (部件哈希集合, 页眉页脚文本集合)：
      - 部件哈希：页眉/页脚部件引用的全部内部部件（图片等）blob 的 SHA-256，
        与 main 的“页眉页脚+其内部图片整体搬迁”对称，防止印章/图片在页眉中丢失；
      - 文本集合：每个生效页眉/页脚部件归一化文本，防止页眉文字错漏
        （main 重写页眉 XML 只改关系/样式 id，不改文字，故比文本而非 XML 字节）。

    只统计本节真正会渲染的类型：default 恒生效；first 需 sectPr 带 w:titlePg；
    even 需 settings 开 w:evenAndOddHeaders。残留但不渲染的模板引用必须排除，
    否则会把模板首页页眉误算成成品附页页眉造成假报警。
    """
    import hashlib
    sect = doc.element.body.find(qn('w:sectPr'))
    title_pg = sect is not None and sect.find(qn('w:titlePg')) is not None
    even_on = doc.settings.element.find(qn('w:evenAndOddHeaders')) is not None
    active_types = {'default'}
    if title_pg:
        active_types.add('first')
    if even_on:
        active_types.add('even')

    part_hashes, texts = set(), set()
    for (kind, wtype), rid in _last_section_effective_refs(doc).items():
        if wtype not in active_types:
            continue
        if not rid or rid not in doc.part.rels:
            continue
        rel = doc.part.rels[rid]
        if rel.is_external:
            continue
        part = rel.target_part
        try:
            xml_root = etree.fromstring(part.blob)
            # 与 main 对称：附页页码应需求取消，PAGE 域文本不计入比对
            _strip_page_number_fields(xml_root)
            texts.add(_norm("".join(t.text or "" for t in xml_root.iter(qn('w:t')))))
        except (etree.XMLSyntaxError, ValueError):
            texts.add(_norm(part.blob.decode('utf-8', 'ignore')))
        for rr in part.rels.values():
            if rr.is_external or rr.target_part is part:
                continue
            part_hashes.add(hashlib.sha256(rr.target_part.blob).hexdigest())
    return part_hashes, texts


# noinspection PyProtectedMember
def _appendix_has_numbering(doc) -> bool:
    """附页边界内是否使用自动编号（w:numPr）。

    自动编号由 numbering.xml 部件定义，不在正文文本里，文本/图片比对都无法覆盖；
    一旦源附页使用自动编号，提示必须走 verify.py 渲染闸门做像素级复核。
    """
    kids = list(doc.element.body)
    b = _apd_boundary(kids)
    if b is None:
        return False
    start, end = b
    for j in range(start, end):
        if kids[j].findall('.//' + qn('w:numPr')):
            return True
    return False


def extract_report(report_path: str) -> Dict[str, Any]:
    doc = Document(report_path)
    tables, paras = _build_doc_index(doc)
    data = {}

    for text in paras:
        text = text.strip()
        if "报告编号" in text:
            match = re.search(r'报告编号[：:]?\s*([^\s]+)', text)
            if match:
                num = match.group(1)
                if len(num) > len(data.get("报告编号", "")):
                    data["报告编号"] = num

    # 基本信息表（机动车号牌/使用单位/道路运输证号/单位地址）
    basic = _find_tbl(tables, ["道路运输证号"])
    if basic is not None:
        data["机动车号牌"] = _norm(_gt(basic, 0, 1))
        data["使用单位"] = _gt(basic, 2, 1)
        data["道路运输证号"] = _norm(_gt(basic, 2, 4))
        data["单位地址"] = _gt(basic, 3, 1)

    # 罐体基本资料表（制造/材质/厚度/装运介质/分仓数量 等）
    base = _find_tbl(tables, ["罐体材质", "分仓数量"])
    if base is not None:
        data["罐体编号"] = _norm(_gt(base, 0, 5))
        data["制造日期"] = _norm(_gt(base, 1, 1))
        data["VIN码"] = _norm(_gt(base, 1, 5))
        data["制造企业"] = _gt(base, 2, 1)
        data["设计代码"] = _norm(_gt(base, 3, 1))
        data["产品标准"] = _norm(_gt(base, 3, 5))
        data["产品型号"] = _norm(_gt(base, 4, 1))
        data["罐体容积"] = _norm(_gt(base, 4, 5))
        data["罐体外形尺寸"] = _norm(_gt(base, 5, 1))
        data["分仓数量"] = _gt(base, 5, 5)
        data["总质量"] = _norm(_gt(base, 6, 1))
        data["核定载质量"] = _norm(_gt(base, 6, 5))
        data["封头材质"] = _gt(base, 8, 2)
        data["筒体材质"] = _gt(base, 9, 2)
        data["封头厚度"] = _norm(_gt(base, 8, 6))
        data["筒体厚度"] = _norm(_gt(base, 9, 6))
        data["装运介质"] = _gt(base, 10, 1)

    # 适装介质列表
    medium_table = _find_tbl(tables, ["介质名称"])
    if medium_table is not None:
        data["适装介质列表"] = _extract_medium_list(medium_table)

    # 测厚记录
    thickness_table = _find_tbl(tables, ["测厚记录", "测厚点部位图"])
    if thickness_table is not None:
        data.update(_extract_thickness(thickness_table))

    # 检验/审核签名日期 + 下次检验日期（单遍扫描）
    data["检验日期"], data["审核日期"], data["下次检验日期"] = _extract_dates(tables)

    # 附页整节（双侧比对用：文本 + 图片哈希），无条件探测
    data["附页文字"] = _extract_appendix_text(doc)
    data["_附页图哈希"] = _appendix_image_hashes(doc)
    data["_附页页眉部件哈希"], data["_附页页眉文本"] = _appendix_chrome(doc)
    data["_附页自动编号"] = _appendix_has_numbering(doc)

    return data

def extract_output(output_path: str) -> Dict[str, Any]:
    doc = Document(output_path)
    tables, paras = _build_doc_index(doc)
    data = {}

    for text in paras:
        text = text.strip()
        if "记录编号：" in text:
            match = re.search(r'记录编号：\s*([^\s]+)', text)
            if match:
                data["报告编号"] = match.group(1)

    t0 = tables[0]
    data["机动车号牌"] = _norm(_gt(t0, 0, 2))
    data["使用单位"] = _gt(t0, 2, 2)
    data["道路运输证号"] = _norm(_gt(t0, 2, 6))
    data["单位地址"] = _gt(t0, 3, 2)
    data["总质量"] = _norm(_gt(t0, 5, 2))
    data["核定载质量"] = _norm(_gt(t0, 5, 6))
    data["制造企业"] = _gt(t0, 6, 2)
    data["制造日期"] = _norm(_gt(t0, 6, 6))
    data["设计代码"] = _norm(_gt(t0, 7, 2))
    data["产品标准"] = _norm(_gt(t0, 7, 6))
    data["产品型号"] = _norm(_gt(t0, 8, 2))
    data["VIN码"] = _norm(_gt(t0, 8, 6))
    data["罐体编号"] = _norm(_gt(t0, 9, 2))
    data["罐体容积"] = _norm(_gt(t0, 9, 6))
    data["罐体外形尺寸"] = _norm(_gt(t0, 10, 2))
    data["分仓数量"] = _gt(t0, 10, 6)
    data["封头材质"] = _gt(t0, 12, 3)
    data["筒体材质"] = _gt(t0, 13, 3)
    data["封头厚度"] = _norm(_gt(t0, 12, 7))
    data["筒体厚度"] = _norm(_gt(t0, 13, 7))
    data["装运介质"] = _gt(t0, 14, 2)

    t3 = tables[3]
    data["封头1"] = _gt(t3, 3, 3)
    data["封头2"] = _gt(t3, 4, 3)
    data["封头3"] = _gt(t3, 5, 3)
    data["封头4"] = _gt(t3, 6, 3)
    data["封头17"] = _gt(t3, 3, 5)
    data["封头18"] = _gt(t3, 4, 5)
    data["封头19"] = _gt(t3, 5, 5)
    data["封头20"] = _gt(t3, 6, 5)
    data["筒体5"] = _gt(t3, 3, 10)
    data["筒体6"] = _gt(t3, 4, 10)
    data["筒体7"] = _gt(t3, 5, 10)
    data["筒体8"] = _gt(t3, 6, 10)
    data["筒体9"] = _gt(t3, 3, 13)
    data["筒体10"] = _gt(t3, 4, 13)
    data["筒体11"] = _gt(t3, 5, 13)
    data["筒体12"] = _gt(t3, 6, 13)
    data["筒体13"] = _gt(t3, 3, 16)
    data["筒体14"] = _gt(t3, 4, 16)
    data["筒体15"] = _gt(t3, 5, 16)
    data["筒体16"] = _gt(t3, 6, 16)

    # 测厚表内的检验/校核日期
    grid3, ncol3, nrow3 = t3
    for r in range(nrow3):
        for c in range(ncol3):
            t = grid3[r * ncol3 + c].strip()
            if "检验" in t and "年" in t and "校核" not in t:
                match = re.search(r'(\d{4}年\d{1,2}月\d{1,2}日)', t)
                if match:
                    data["检验日期"] = match.group(1)
            if "校核" in t and "年" in t:
                match = re.search(r'(\d{4}年\d{1,2}月\d{1,2}日)', t)
                if match:
                    data["审核日期"] = match.group(1)

    # 提取原始记录里的适装介质列表（第 0 张表 row17 起，行数随表动态扩展）
    try:
        medium_list = []
        max_media_rows = max(0, t0[2] - 17)
        for i in range(max_media_rows):
            r = 17 + i
            if r >= t0[2]:
                break
            seq = _gt(t0, r, 0)
            if not seq.isdigit():
                continue
            name = _gt(t0, r, 1)
            if name in ('', '——'):
                continue
            medium_list.append({
                "序号": seq,
                "介质名称": name,
                "UN号": _gt(t0, r, 4),
                "类别及项别": _gt(t0, r, 5),
                "包装类别": _gt(t0, r, 6),
                "备注": _gt(t0, r, 8),
            })
        data["适装介质列表"] = medium_list
    # 介质区提取容忍任何版式异常，缺数据时由两侧比对/判空兜底
    # noinspection PyBroadException
    except Exception:
        pass

    # 下次检验日期（遍历所有表格，取第一个带日期的命中，与报告侧口径一致；
    # 归一化必须用 _norm 去掉全部空白含全角空格，不能只 replace 半角空格）
    try:
        _nxt_found = False
        for grid, ncol, nrow in tables:
            if _nxt_found:
                break
            for r in range(nrow):
                if _nxt_found:
                    break
                for c in range(ncol):
                    text = grid[r * ncol + c].strip()
                    if "下次检验日期" in text:
                        match = re.search(r'下次检验日期[：:]?\s*(\d{4}\s*年\s*\d{1,2}\s*月)', text)
                        if match:
                            data["下次检验日期"] = _norm(match.group(1))
                            _nxt_found = True
                            break
    # 下次检验日期只是补充字段，提取失败不影响其他比对
    # noinspection PyBroadException
    except Exception:
        pass

    # 充装介质附页整节（无条件探测；文本+图片哈希，与源侧对称比对）
    data["附页文字"] = _extract_appendix_text(doc)
    data["_附页图哈希"] = _appendix_image_hashes(doc)
    data["_附页页眉部件哈希"], data["_附页页眉文本"] = _appendix_chrome(doc)
    data["_附页自动编号"] = _appendix_has_numbering(doc)

    return data


def _safe_extractall(zf: zipfile.ZipFile, dest: str) -> None:
    """带 Zip-Slip 防护的解压：拒绝任何会跳出目标目录的成员路径。
    并兼容 GBK 文件名：未置 UTF-8 标志的 zip（如 Windows 自带压缩）会被 zipfile
    误按 cp437 解码成乱码，这里回退 GBK 重新解码还原中文名。"""
    base = os.path.abspath(dest)
    for info in zf.infolist():
        # 解码后的名字只用于计算落盘路径，绝不能写回 info.filename：
        # Python 3.13 的 zf.open() 会拿该名回查 NameToInfo（键仍是原始名），
        # 改写后会误抛 KeyError。与 main.py 的同名函数保持一致。
        name = info.filename
        # 未置 UTF-8 标志(0x800)时按 GBK 重解码；编解码任一失败则保留原名
        if not (info.flag_bits & 0x800):
            try:
                name = name.encode('cp437').decode('gbk')
            except (UnicodeEncodeError, UnicodeDecodeError):
                pass
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


def _extract_nested_zips(root: str, max_depth: int = 3) -> Tuple[Set[str], List[str]]:
    """递归解包 root 内的嵌套 zip：就地解压到 <zip名去扩展名>/ 子目录。

    返回 (已成功解压的 zip 绝对路径集合, 解压失败的 zip 相对路径名单)。
    深度上限 max_depth 防止恶意/异常的无限嵌套；按绝对路径去重防重复解包。
    .7z/.rar 不解，仍由调用方的压缩包告警统一提示人工核对。
    """
    extracted: Set[str] = set()
    failed: List[str] = []
    frontier = [root]
    for _depth in range(max_depth):
        if not frontier:
            break
        next_frontier = []
        for base in frontier:
            for dirpath, _dirs, files in os.walk(base):
                for f in files:
                    if not f.lower().endswith('.zip'):
                        continue
                    zp = os.path.abspath(os.path.join(dirpath, f))
                    if zp in extracted:
                        continue
                    dest = os.path.join(dirpath, os.path.splitext(f)[0])
                    os.makedirs(dest, exist_ok=True)
                    try:
                        with zipfile.ZipFile(zp, 'r') as zf:
                            _safe_extractall(zf, dest)
                        extracted.add(zp)
                        print(f"  嵌套解压: {os.path.relpath(zp, root)}")
                        next_frontier.append(dest)
                    # 嵌套包解压失败只记录名单（汇总告警），不阻断其他批次
                    # noinspection PyBroadException
                    except Exception as e:
                        print(f"  ！嵌套压缩包解压失败: "
                              f"{os.path.relpath(zp, root)} ({e})")
                        failed.append(os.path.relpath(zp, root))
        frontier = next_frontier
    return extracted, failed


def _ensure_extracted(report_folder: str,
                      extract_tag: str = "_已解压_check") -> Tuple[Optional[str], List[str]]:
    """扫描 zip 并解压，返回 (解压后根目录或None, 损坏zip名单)。

    任何一个 zip 解压失败都计入名单返回——调用方必须据此 fail-closed，
    不能在 zip 损坏时给出“全部通过”的假象。
    """
    extract_root = os.path.join(report_folder, extract_tag)
    failed = []
    if not os.path.exists(report_folder):
        return None, failed

    # 非 zip 压缩包（.7z/.rar）任何时刻都不自动解包，显式告警提醒人工核对份数
    other_archives = [f for f in os.listdir(report_folder)
                      if f.lower().endswith(('.7z', '.rar'))
                      and os.path.isfile(os.path.join(report_folder, f))]
    if other_archives:
        print(f"  ！发现 {len(other_archives)} 个非 zip 压缩包（.7z/.rar），程序不会自动解包，"
              f"请人工解压或核对份数:")
        for a in sorted(other_archives):
            print(f"      {a}")

    zip_files = [f for f in os.listdir(report_folder)
                 if f.lower().endswith('.zip') and os.path.isfile(os.path.join(report_folder, f))]
    if not zip_files:
        return None, failed

    print(f"检测到 {len(zip_files)} 个压缩包，自动解压用于匹配...")
    if os.path.exists(extract_root):
        shutil.rmtree(extract_root, ignore_errors=True)
    os.makedirs(extract_root, exist_ok=True)

    # 每个 zip 单独解压到 _已解压_check/<zip名>/，与 main 的解压结构一致，
    # 也避免多个 zip 的散文件混放、同名互相覆盖
    for zname in zip_files:
        zpath = os.path.join(report_folder, zname)
        target = os.path.join(extract_root, os.path.splitext(zname)[0])
        os.makedirs(target, exist_ok=True)
        try:
            with zipfile.ZipFile(zpath, 'r') as zf:
                _safe_extractall(zf, target)
            print(f"  已解压: {zname}")
        # 解压失败的 zip 必须原样计入名单，因此这里捕获全部异常
        # noinspection PyBroadException
        except Exception as e:
            print(f"  解压失败 {zname}: {e}")
            failed.append(zname)

    # 自动递归解包嵌套 zip（深度上限 3），与 main 口径一致；
    # 嵌套包解压失败 = 源数据不完整，必须计入 failed 让调用方 fail-closed
    nested_ok: Set[str] = set()
    for zentry in sorted(os.listdir(extract_root)):
        zfull = os.path.join(extract_root, zentry)
        if not os.path.isdir(zfull):
            continue
        ok, bad = _extract_nested_zips(zfull)
        nested_ok |= ok
        failed.extend(bad)

    # 已自动解包的嵌套 zip 不再告警；只点名 .7z/.rar 及超深度/解压失败的 zip，
    # 这些源数据可能不完整，必须显式告警提醒人工核对份数
    stray = []
    for root, _dirs, files in os.walk(extract_root):
        for f in files:
            low = f.lower()
            if not low.endswith(('.zip', '.7z', '.rar')):
                continue
            fp = os.path.abspath(os.path.join(root, f))
            if low.endswith('.zip') and fp in nested_ok:
                continue
            stray.append(os.path.relpath(fp, extract_root))
    if stray:
        print(f"  ！发现 {len(stray)} 个无法自动解包的压缩包"
              f"（.7z/.rar 或超嵌套深度/解压失败的 zip），请人工核对份数:")
        for a in sorted(stray):
            print(f"      {a}")
    return extract_root, failed


def _convert_non_docx_reports(search_root: str) -> None:
    """递归把目录下 .wps/.doc/.docm 检验报告转成 .docx（与 main.py 口径一致）。

    同名 .docx 已存在则跳过。转换产物在临时解压目录里，用完随目录一起清理。
    """
    valid_exts = ('.wps', '.doc', '.docm')
    candidates = []
    for root, dirs, files in os.walk(search_root):
        for f in files:
            if f.startswith('~$') or '合格证' in f:
                continue
            if f.lower().endswith(valid_exts):
                src = os.path.join(root, f)
                dst = os.path.splitext(src)[0] + ".docx"
                if not os.path.exists(dst):
                    candidates.append((src, dst))
    if not candidates:
        return

    print(f"检测到 {len(candidates)} 份非 .docx 检验报告，正在用 WPS 转换...")
    word = None
    try:
        import win32com.client
        word = win32com.client.Dispatch("KWps.Application")
        word.Visible = False
        # COM 属性设置可能不被支持，忽略即可
        # noinspection PyBroadException
        try:
            word.DisplayAlerts = 0  # wdAlertsNone：损坏/加密文件不弹模态框卡死整批
        except Exception:
            pass
        converted = 0
        for src, dst in candidates:
            try:
                doc = word.Documents.Open(os.path.abspath(src))
                doc.SaveAs2(os.path.abspath(dst), FileFormat=16)  # 16 = wdFormatXMLDocument
                doc.Close()
                converted += 1
                print(f"  已转换: {os.path.basename(src)}")
            # COM 单文件失败不影响其他文件转换
            # noinspection PyBroadException
            except Exception as e:
                print(f"  转换失败 {os.path.basename(src)}: {e}")
        print(f"转换完成：成功 {converted}/{len(candidates)}")
    except ImportError:
        print("未安装 pywin32 或 WPS，无法转换 .wps/.doc 报告，相关文件将匹配失败")
    # WPS 启动/整体异常：COM 异常类型不可控，统一兜底
    # noinspection PyBroadException
    except Exception as e:
        print(f"WPS 转换过程出错: {e}")
    finally:
        if word is not None:
            # noinspection PyBroadException
            try:
                word.Quit()
            except Exception:
                pass


def _prune_ignored_dirs(dirs: List[str]) -> None:
    """os.walk 原地剪枝：下划线临时目录（_已解压/_已解压_check 残留）、
    __pycache__、成品目录等一律不下钻，避免同一报告被索引两次造成假“同名歧义”。
    口径与 main.py 的批次收集（下划线目录直接忽略）对齐。
    """
    dirs[:] = [d for d in dirs
               if not d.startswith('_')
               and d not in ('__pycache__', '原始记录汇总')]


def _match_key(filename: str) -> str:
    """把源/成品文件名归一化为「报告编号+号牌」匹配键。

    成品命名规则（main.py）：报告编号+号牌+原始记录.docx；
    历史源文件名尾部可能带「检验报告」「报告」，还可能有「 - 副本」或多余空格。
    剥离这些尾部词后，源与成品按同一个键匹配，不再依赖完整文件名是否一致。
    """
    stem = os.path.splitext(os.path.basename(filename))[0]
    stem = re.sub(r'\s*-\s*副本\s*$', '', stem)
    # 循环剥离：旧成品名尾部可能叠着多个词，如「...检验报告原始记录」
    prev = None
    while prev != stem:
        prev = stem
        stem = re.sub(r'(?:原始记录|检验报告|报告)\s*$', '', stem)
    return stem.strip()


def _build_report_index(*roots: str) -> Dict[str, List[str]]:
    """一次性遍历目录，建立 文件名/匹配键 ->[完整路径...] 索引。

    同时按完整文件名与「编号+号牌」匹配键建索引：
      - 完整名键保留（成品可能与源完全同名的情形）；
      - 匹配键把「...检验报告.docx」「...报告.docx」「...原始记录.docx」
        统一到「编号+号牌」，源文件名尾词不同也能命中。
    同键收集为列表（不再先到先得静默吞歧义），调用方按长度决定唯一/歧义。
    """
    index = {}
    for base in roots:
        if not base or not os.path.isdir(base):
            continue
        for root, dirs, files in os.walk(base):
            _prune_ignored_dirs(dirs)
            for f in files:
                full_path = os.path.join(root, f)
                index.setdefault(f, []).append(full_path)
                if _is_report_input_name(f):
                    key = _match_key(f)
                    if key and key != f:
                        index.setdefault(key, []).append(full_path)
    return index


def _is_report_input_name(filename: str) -> bool:
    """识别报告输入文件名（.docx/.doc/.wps/.docm，排除锁文件与合格证）。"""
    if filename.startswith("~$"):
        return False
    if "合格证" in filename:
        return False
    return filename.lower().endswith((".docx", ".doc", ".wps", ".docm"))


def _direct_report_stems(dirpath: str) -> Set[str]:
    """目录直接层报告的「编号+号牌」匹配键集合（不下钻）。"""
    try:
        return {_match_key(f) for f in os.listdir(dirpath)
                if os.path.isfile(os.path.join(dirpath, f)) and _is_report_input_name(f)}
    except OSError:
        return set()


def _recursive_report_stems(dirpath: str) -> Set[str]:
    """目录整树报告的「编号+号牌」匹配键集合（剪枝忽略目录）。"""
    stems = set()
    for r, dirs, files in os.walk(dirpath):
        _prune_ignored_dirs(dirs)
        for f in files:
            if _is_report_input_name(f):
                stems.add(_match_key(f))
    return stems


def _visible_subdirs(dirpath: str) -> List[str]:
    """直接子目录（排除下划线临时目录与内部目录），与 main 的批次收集口径一致。"""
    try:
        return sorted(d for d in os.listdir(dirpath)
                      if os.path.isdir(os.path.join(dirpath, d))
                      and not d.startswith('_')
                      and d not in ('__pycache__', '原始记录汇总'))
    except OSError:
        return []


def _scan_source_batches(root: str) -> Dict[str, Set[str]]:
    """检验报告文件夹口径（与 main 方式2/3 完全一致）：
    root 直接层散文件 → 批次名 = root 目录名（仅直接层）；
    每个直接子目录 → 各自递归归并为一个批次（批次名 = 子目录名，含更深层文件）。
    """
    batches = {}
    if not root or not os.path.isdir(root):
        return batches
    root_stems = _direct_report_stems(root)
    if root_stems:
        batches[os.path.basename(root.rstrip('\\/'))] = root_stems
    for d in _visible_subdirs(root):
        stems = _recursive_report_stems(os.path.join(root, d))
        if stems:
            batches.setdefault(d, set()).update(stems)
    return batches


def _scan_zip_batches(extract_root: str) -> Dict[str, Set[str]]:
    """zip 解压目录口径（与 main 方式1 完全一致）：
    每个 zipstem 的内层目录各自递归归并为一个批次；
    zipstem 无内层目录时其直接层散文件为 zipstem 名批次；
    zipstem 有内层目录且自身还散放报告时，这些散文件 main 不会处理——
    单列成 zipstem 名批次，让双向盘点能把这批被丢弃的报告逐份点名。
    """
    batches = {}
    if not extract_root or not os.path.isdir(extract_root):
        return batches
    for z in _visible_subdirs(extract_root):
        zfull = os.path.join(extract_root, z)
        inner = _visible_subdirs(zfull)
        if inner:
            for d in inner:
                stems = _recursive_report_stems(os.path.join(zfull, d))
                if stems:
                    batches.setdefault(d, set()).update(stems)
        # 无论有无内层目录，zipstem 直接层散文件都单列（main 只在无内层目录时才收它们；
        # 有内层目录时 main 丢弃这些散文件，这里单列后盘点会报"源有、成品缺失"）
        direct = _direct_report_stems(zfull)
        if direct:
            batches.setdefault(z, set()).update(direct)
    return batches


def _scan_output_batches(output_folder: str) -> Dict[str, Set[str]]:
    """扫描成品目录，返回 {批次名: set(成品stem)}。成品目录名以“原始记录”结尾。"""
    batches = {}
    if not os.path.isdir(output_folder):
        return batches
    for d in os.listdir(output_folder):
        full = os.path.join(output_folder, d)
        if not (os.path.isdir(full) and d.endswith("原始记录")):
            continue
        name = d[:-4]
        stems = set()
        for root, _d, files in os.walk(full):
            for f in files:
                if f.lower().endswith(".docx") and not f.startswith("~$"):
                    # 统一归一化为「编号+号牌」匹配键，与源侧口径一致
                    stems.add(_match_key(f))
        batches[name] = stems
    return batches


def _reconcile_batches(src_maps: Iterable[Dict[str, Set[str]]],
                       output_folder: str) -> List[str]:
    """批次级双向盘点：源↔成品的批次与文件必须一一对应，任一缺失逐条列出。

    入参为多张源侧批次表（检验报告文件夹口径 + zip 解压目录口径），先合并再比对。
    """
    src: Dict[str, Set[str]] = {}
    for m in src_maps:
        for k, v in m.items():
            src.setdefault(k, set()).update(v)
    out = _scan_output_batches(output_folder)

    issues = []
    for name in sorted(set(src) | set(out)):
        s, o = src.get(name), out.get(name)
        if s is None:
            issues.append("批次「%s」：成品有 %d 份，但无源批次目录" % (name, len(o)))
            continue
        if o is None:
            issues.append("批次「%s」：源有 %d 份，但无成品目录" % (name, len(s)))
            continue
        for f in sorted(s - o):
            issues.append("批次「%s」：源报告有、成品缺失：%s" % (name, f))
        for f in sorted(o - s):
            issues.append("批次「%s」：成品有、源缺失：%s" % (name, f))
    return issues


def _format_diff_detail(diff_fields: List[Tuple[str, Any, Any]]) -> List[str]:
    """把字段差异列表渲染成末尾汇总用的文本行。"""
    detail = []
    for field, r_val, o_val in diff_fields:
        detail.append("    - " + field + ":")
        if isinstance(r_val, list):
            detail.append("      检验报告介质数=%d" % len(r_val))
            for item in r_val:
                detail.append("        %s. %s  UN=%s  类别=%s  包装=%s" % (
                    item.get("序号", ""), item.get("介质名称", ""),
                    item.get("UN号", ""), item.get("类别及项别", ""), item.get("包装类别", "")))
            detail.append("      原始记录介质数=%d" % len(o_val))
            for item in o_val:
                detail.append("        %s. %s  UN=%s  类别=%s  包装=%s" % (
                    item.get("序号", ""), item.get("介质名称", ""),
                    item.get("UN号", ""), item.get("类别及项别", ""), item.get("包装类别", "")))
        else:
            r_display = r_val if r_val else "(空)"
            o_display = o_val if o_val else "(空)"
            detail.append("      原始记录=[" + o_display + "]")
            detail.append("      检验报告=[" + r_display + "]")
    return detail


def _extract_xml_identity(docx_path: str) -> Dict[str, str]:
    """异构通道：直接读 zip 内 word/document.xml 原始 XML（不经 python-docx
    的表格网格/合并单元格重建逻辑），独立提取报告编号与机动车号牌。

    与主通道构成两套独立实现，避免共享同一组解析假设。
    """
    with zipfile.ZipFile(docx_path) as z:
        xml_root = etree.fromstring(z.read("word/document.xml"))

    def wt_text(kid):
        return "".join(t.text or "" for t in kid.iter(qn("w:t")))

    # 报告编号（成品用“记录编号”，源报告用“报告编号”，两者都认，取最长）
    num = ""
    for m in re.finditer(r"(?:报告编号|记录编号)[：:]?\s*(B-[A-Za-z0-9\-]+)",
                         wt_text(xml_root)):
        if len(m.group(1)) > len(num):
            num = m.group(1)

    # 号牌：含“道路运输证号”的表，首行第一个非“号牌”标签的非空单元格
    plate = ""
    for tbl in xml_root.iter(qn("w:tbl")):
        if "道路运输证号" in wt_text(tbl):
            first_tr = tbl.find(qn("w:tr"))
            if first_tr is not None:
                for tc in first_tr.findall(qn("w:tc")):
                    tx = _norm(wt_text(tc))
                    if tx and "号牌" not in tx:
                        plate = tx
                        break
            break

    # 介质表有效数据行数（独立于 python-docx 解析器的第三通道计数）。
    # 用于发现“表头版式未识别→整表介质静默丢失”这类双解析器同错绿灯的盲区。
    media_rows = 0
    for tbl in xml_root.iter(qn("w:tbl")):
        txt = wt_text(tbl)
        if "介质名称" in txt and "序号" in txt:
            for tr in tbl.findall(qn("w:tr")):
                tcs = tr.findall(qn("w:tc"))
                if len(tcs) > 1:
                    if _norm(wt_text(tcs[0])).isdigit() \
                            and _norm(wt_text(tcs[1])) not in ('', '——'):
                        media_rows += 1
    return {"报告编号": num, "机动车号牌": plate, "_介质行数": media_rows}


def _check_one(task: Tuple[str, str, Optional[str]]) -> Tuple[str, str, int, List[str]]:
    """多进程 worker：只做纯 python-docx 读取与字段比对（不含 COM/打印）。

    task = (output_path, rel_path, report_path_or_空_or_None)
      report_path = None 表示匹配到多份同名源报告（歧义，fail-closed）
    返回 (rel_path, status, 异常数, 明细行列表)；status ∈ {'ok','notfound','error'}
    """
    output_path, rel_path, report_path = task
    report_key = _match_key(output_path)

    if report_path is None:
        return (rel_path, "error", 1,
                ["  匹配到多份同名检验报告，无法唯一确定（请人工核对批次目录）"])
    if not report_path:
        return (rel_path, "notfound", 1,
                ["  未找到对应的检验报告 [%s]，已跳过" % report_key])
    try:
        report_data = extract_report(report_path)
        output_data = extract_output(output_path)

        diff_fields = []
        for field in FIELDS:
            r_val = report_data.get(field, "").strip()
            o_val = output_data.get(field, "").strip()
            if r_val != o_val:
                diff_fields.append((field, r_val, o_val))

        r_list = report_data.get("适装介质列表", [])
        o_list = output_data.get("适装介质列表", [])
        if r_list != o_list:
            diff_fields.append(("适装介质列表", r_list, o_list))

        # 源头判空：报告侧必填字段为空时，记录侧必然同步空白，两边比对查不出来
        missing_fields = []
        for field in REQUIRED_FIELDS:
            val = report_data.get(field, "")
            if isinstance(val, (list, tuple, dict)):
                is_empty = len(val) == 0
            else:
                is_empty = not str(val).strip()
            if is_empty:
                missing_fields.append(field)

        # 多余空格：报告原文中的连续空格/全角空格会原样拷进记录，比对同样查不出来
        space_fields = []
        for field in _SPACE_CHECK_FIELDS:
            val = report_data.get(field, "")
            if isinstance(val, str) and val.strip() and _SPACE_RE.search(val):
                space_fields.append((field, val))

        # 充装介质附页：双侧文本 + 图片哈希严格比对（不再只看“有没有”）
        apd_issues = []
        r_apd = report_data.get("附页文字", "")
        o_apd = output_data.get("附页文字", "")
        if r_apd != o_apd:
            if r_apd and not o_apd:
                apd_issues.append("原始记录缺少充装介质附页（源报告有附页，未追加或追加失败）")
            elif o_apd and not r_apd:
                apd_issues.append("原始记录多出源报告不存在的附页内容")
            else:
                apd_issues.append("附页文本双侧不一致（源 %d 字，成品 %d 字，请回源比对）"
                                  % (len(r_apd), len(o_apd)))
        r_imgs = report_data.get("_附页图哈希", set())
        o_imgs = output_data.get("_附页图哈希", set())
        if r_imgs != o_imgs:
            apd_issues.append("附页图片（印章等）不一致：源 %d 张、成品 %d 张，"
                              "可能存在图片丢失或替换" % (len(r_imgs), len(o_imgs)))
        # 页眉页脚/自动编号检查只在确有附页节时进行：普通报告的正文节页眉与
        # 模板页眉天然不同，无条件比较会造成全量假报警
        if r_apd or o_apd:
            # 附页末节页眉/页脚：内部图片部件必须一一对应（防页眉印章丢失）
            r_chrome_img = report_data.get("_附页页眉部件哈希", set())
            o_chrome_img = output_data.get("_附页页眉部件哈希", set())
            if r_chrome_img != o_chrome_img:
                apd_issues.append("附页页眉/页脚内图片不一致：源 %d 个、成品 %d 个，"
                                  "可能存在页眉印章/图片丢失" % (len(r_chrome_img), len(o_chrome_img)))
            # 附页末节页眉/页脚文字必须一致（页眉文字错漏在正文比对中不可见）
            r_chrome_txt = report_data.get("_附页页眉文本", set())
            o_chrome_txt = output_data.get("_附页页眉文本", set())
            if r_chrome_txt != o_chrome_txt:
                apd_issues.append("附页页眉/页脚文字不一致：源 %d 段、成品 %d 段"
                                  % (len(r_chrome_txt), len(o_chrome_txt)))
            # 自动编号定义无法经文本比对核对：提示必须走 verify.py 渲染闸门
            if report_data.get("_附页自动编号") or output_data.get("_附页自动编号"):
                apd_issues.append("附页使用了自动编号（numbering），L2 无法核对编号定义，"
                                  "请运行 verify.py 做渲染像素复核")
        # 源报告自相矛盾：装运栏声明见附页，但正文找不到附页
        zy_r = _norm(report_data.get("装运介质", ""))
        claims_appendix = ("附" in zy_r) or bool(re.search(r"末页|后页|清单", zy_r))
        if claims_appendix and not r_apd and not o_apd:
            apd_issues.append("装运介质写明见附页/末页等，但源报告未找到附页正文，请人工核查")

        # 异构通道：原始 XML 独立实现 vs python-docx 通道，标识字段必须三方一致
        hetero = []
        try:
            xr = _extract_xml_identity(report_path)
            xo = _extract_xml_identity(output_path)
            for f in ("报告编号", "机动车号牌"):
                a, b = _norm(xr.get(f)), _norm(xo.get(f))
                pr, po = _norm(report_data.get(f)), _norm(output_data.get(f))
                if a and a != pr:
                    hetero.append("源报告 %s 两通道分歧：原始XML=[%s]，docx提取=[%s]" % (f, a, pr))
                if b and b != po:
                    hetero.append("成品 %s 两通道分歧：原始XML=[%s]，docx提取=[%s]" % (f, b, po))
                if a and b and a != b:
                    hetero.append("%s 双侧不一致：源=[%s]，成品=[%s]" % (f, a, b))
            # 共享解析器盲区探针1：介质表有有效数据行但解析结果为空
            # （表头版式未识别→整表介质静默丢失，双解析器会同错绿灯）
            if xr.get("_介质行数", 0) > 0 and not report_data.get("适装介质列表"):
                hetero.append("源报告介质表存在 %d 行有效数据但解析结果为空（表头版式未识别），"
                              "适装介质整表未拷贝" % xr["_介质行数"])
            # 探针2：测厚点超出模板20点容量（解析器收得到但模板写不下，字段比对看不见）
            extra_pts = sorted(k for k in report_data
                               if re.fullmatch(r'(?:封头|筒体)\d+', k) and k not in FIELDS)
            if extra_pts:
                hetero.append("源报告测厚点超出模板20点容量（%s），这些点不会写入成品"
                              % "、".join(extra_pts))
        # 异构通道本身就是独立探针，任何失败都要落成问题记录，不能抛出
        # noinspection PyBroadException
        except Exception as e:
            hetero.append("异构通道执行出错: %s" % str(e)[:50])

        if not diff_fields and not missing_fields and not space_fields \
                and not apd_issues and not hetero:
            return (rel_path, "ok", 0, [])

        detail = []
        total = 0
        if missing_fields:
            detail.append("  检验报告漏填字段（原始记录同步为空，比对无法发现），共 %d 项:" % len(missing_fields))
            for field in missing_fields:
                detail.append("    - " + field)
            total += len(missing_fields)
        if space_fields:
            detail.append("  检验报告字段含多余空格（已原样拷贝，请回报告核对），共 %d 项:" % len(space_fields))
            for field, val in space_fields:
                detail.append("    - %s: [%s]" % (field, val))
            total += len(space_fields)
        if apd_issues:
            detail.append("  充装介质附页（双侧比对），共 %d 项:" % len(apd_issues))
            for x in apd_issues:
                detail.append("    - " + x)
            total += len(apd_issues)
        if hetero:
            detail.append("  异构通道（原始XML独立核对），共 %d 项:" % len(hetero))
            for x in hetero:
                detail.append("    - " + x)
            total += len(hetero)
        if diff_fields:
            detail.append("  检验报告与原始记录不一致，共 %d 项:" % len(diff_fields))
            detail.extend(_format_diff_detail(diff_fields))
            total += len(diff_fields)
        return (rel_path, "error", total, detail)
    # worker 边界必须兜住一切异常并落成错误记录，不能让进程池吞掉
    # noinspection PyBroadException
    except Exception as e:
        return (rel_path, "error", 1, ["  检查出错: " + str(e)[:50]])


# ============================================================
# 抽样人工复核：每批次随机抽 5%（至少 1 份），生成 HTML 供肉眼复核
# ============================================================

REVIEW_FOLDER = "抽样复核"

_REVIEW_CSS = """
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:"Microsoft YaHei","PingFang SC",sans-serif;background:#f5f6f8;color:#222}
#sidebar{position:fixed;left:0;top:0;bottom:0;width:240px;overflow-y:auto;
background:#2c3e50;color:#ecf0f1;padding:16px 12px;z-index:10}
#sidebar h1{font-size:16px;color:#fff;margin-bottom:10px;border-bottom:1px solid #4a6278;padding-bottom:8px}
#sidebar .batch{margin:8px 0 4px;font-size:13px;color:#f1c40f;font-weight:bold}
#sidebar a{display:block;color:#cfd8dc;text-decoration:none;font-size:12px;
padding:4px 6px;border-radius:4px;word-break:break-all;line-height:1.4}
#sidebar a:hover{background:#3d566e;color:#fff}
#main{margin-left:240px;padding:24px 16px;display:flex;flex-direction:column;align-items:center}
.card{width:100%;max-width:960px;background:#fff;border-radius:8px;margin-bottom:28px;
box-shadow:0 1px 4px rgba(0,0,0,.12);overflow:hidden}
.card h2{font-size:15px;padding:12px 16px;background:#34495e;color:#fff;word-break:break-all}
.card .paths{padding:8px 16px;font-size:12px;color:#666;background:#ecf0f1;word-break:break-all;line-height:1.6}
.card table{width:100%;border-collapse:collapse;font-size:13px}
.card th,.card td{border:1px solid #dde3e8;padding:6px 8px;text-align:left;
word-break:break-all;vertical-align:top}
.card th{background:#f0f3f6;font-weight:bold}
.card td.fn{width:96px;background:#fafbfc;font-weight:bold}
tr.bad td{background:#fdecea;color:#c0392b;font-weight:bold}
tr.bad td:first-child{color:#c0392b}
.warn{padding:14px 16px;color:#c0392b;font-weight:bold}
#summary{width:100%;max-width:960px;background:#fff;border-radius:8px;padding:16px 20px;
margin-bottom:24px;box-shadow:0 1px 4px rgba(0,0,0,.12);font-size:14px;line-height:1.9}
#top{scroll-margin-top:0}
@media (max-width:768px){
#sidebar{position:static;width:100%;max-height:none}
#main{margin-left:0;padding:12px 8px}
.card h2{font-size:14px}
.card th,.card td{padding:5px 6px;font-size:12px}
.card td.fn{width:72px}
}
"""


def _review_esc(s: str) -> str:
    return html.escape(s, quote=True)


def _collect_outputs(output_folder: str,
                     batch_filter: Optional[str] = None) -> Dict[str, List[str]]:
    """按批次（第一级目录名）归组所有成品 docx。"""
    groups: Dict[str, List[str]] = {}
    for walk_root, _dirs, files in os.walk(output_folder):
        for f in files:
            if not (f.lower().endswith(".docx") and not f.startswith("~$")):
                continue
            rel = os.path.relpath(os.path.join(walk_root, f), output_folder)
            batch = rel.split(os.sep)[0]
            groups.setdefault(batch, []).append(os.path.join(walk_root, f))
    if batch_filter:
        prefix = batch_filter + "原始记录"
        groups = {b: v for b, v in groups.items() if b == prefix}
    for v in groups.values():
        v.sort()
    return dict(sorted(groups.items()))


def _compare_pair(src_path: Optional[str], out_path: str
                  ) -> Tuple[List[Tuple[str, str, str, bool]],
                             List[Tuple[str, str, bool]], Optional[str]]:
    """提取并比对一对（源报告, 成品）。
    返回 (字段行列表(字段,源值,成品值,是否一致), 介质行列表(源介质,成品介质,是否一致), 错误信息或None)。
    """
    if not src_path:
        return [], [], "未找到唯一对应的源报告（缺失或同名歧义），无法比对"
    try:
        src = extract_report(src_path)
        out = extract_output(out_path)
    # 单份提取失败只标记该卡片，不中断整批
    # noinspection PyBroadException
    except Exception as e:
        return [], [], "字段提取失败: %s" % str(e)[:80]

    rows = []
    for f in FIELDS:
        sv = str(src.get(f, "") or "")
        ov = str(out.get(f, "") or "")
        rows.append((f, sv, ov, _norm(sv) == _norm(ov)))

    sm = [r.get("介质名称", "") for r in src.get("适装介质列表", [])]
    om = [r.get("介质名称", "") for r in out.get("适装介质列表", [])]
    media_rows = []
    for i in range(max(len(sm), len(om))):
        s = sm[i] if i < len(sm) else ""
        o = om[i] if i < len(om) else ""
        media_rows.append((s, o, _norm(s) == _norm(o)))
    return rows, media_rows, None


def _render_review_html(sampled, results, percent: float, seed: int) -> str:
    """sampled: [(批次, 成品路径, 源路径或None, 锚点)]；results: {成品路径: _compare_pair结果}"""
    parts = ["<!DOCTYPE html><html lang=\"zh-CN\"><head><meta charset=\"utf-8\">",
             "<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">",
             "<title>抽样人工复核单</title><style>", _REVIEW_CSS, "</style></head><body>"]

    parts.append("<div id=\"sidebar\"><h1>抽样复核目录</h1>")
    parts.append("<a href=\"#top\">▲ 回到顶部（汇总）</a>")
    cur_batch = None
    for batch, op, _src, anchor in sampled:
        if batch != cur_batch:
            cur_batch = batch
            parts.append("<div class=\"batch\">%s</div>" % _review_esc(batch))
        parts.append("<a href=\"#%s\">%s</a>" % (anchor, _review_esc(os.path.basename(op))))
    parts.append("</div>")

    parts.append("<div id=\"main\"><div id=\"top\"></div>")

    total = len(sampled)
    n_bad = sum(1 for _b, op, _s, _a in sampled
                if results[op][2] or any(not m for *_x, m in results[op][0])
                or any(not m for *_x, m in results[op][1]))
    parts.append("<div id=\"summary\">")
    parts.append("<b>抽样人工复核单</b><br>")
    parts.append("生成时间：%s<br>" % time.strftime("%Y-%m-%d %H:%M:%S"))
    parts.append("抽样比例：每批次 %g%%（至少 1 份），随机种子：%d<br>" % (percent, seed))
    parts.append("抽样总数：%d 份；<span style=\"color:%s;font-weight:bold\">"
                 "发现不一致：%d 份</span><br>"
                 % (total, "#c0392b" if n_bad else "#27ae60", n_bad))
    parts.append("复核方法：逐行对照“源报告”与“原始记录”两列，红色行为不一致项，"
                 "须打开原文件人工核实。")
    parts.append("</div>")

    for batch, op, src, anchor in sampled:
        rows, media_rows, err = results[op]
        parts.append("<div class=\"card\" id=\"%s\">" % anchor)
        parts.append("<h2>%s　<small style=\"font-weight:normal\">[%s]</small></h2>"
                     % (_review_esc(os.path.basename(op)), _review_esc(batch)))
        parts.append("<div class=\"paths\">成品：%s<br>源报告：%s</div>"
                     % (_review_esc(op), _review_esc(src or "（未找到）")))
        if err:
            parts.append("<div class=\"warn\">！%s</div>" % _review_esc(err))
        else:
            parts.append("<table><tr><th style=\"width:96px\">字段</th>"
                         "<th>源报告</th><th>原始记录</th></tr>")
            for f, sv, ov, match in rows:
                cls = "" if match else " class=\"bad\""
                parts.append("<tr%s><td class=\"fn\">%s</td><td>%s</td><td>%s</td></tr>"
                             % (cls, _review_esc(f), _review_esc(sv), _review_esc(ov)))
            if media_rows:
                parts.append("<tr><th colspan=\"3\" style=\"background:#e8eef4\">"
                             "适装介质列表（介质名称逐行对照）</th></tr>")
                for s, o, match in media_rows:
                    cls = "" if match else " class=\"bad\""
                    parts.append("<tr%s><td class=\"fn\">介质</td><td>%s</td><td>%s</td></tr>"
                                 % (cls, _review_esc(s), _review_esc(o)))
            parts.append("</table>")
        parts.append("</div>")

    parts.append("</div></body></html>")
    return "".join(parts)


def _generate_sample_review(output_folder: str, report_index: Dict[str, List[str]],
                            percent: float = 5.0,
                            seed: Optional[int] = None) -> Optional[str]:
    """生成抽样复核 HTML，返回文件路径（无成品时返回 None）。"""
    groups = _collect_outputs(output_folder)
    if not groups:
        return None

    if seed is None:
        seed = int(time.time())
    rng = random.Random(seed)
    sampled = []
    seq = 0
    for batch, outs in groups.items():
        n = max(1, math.ceil(len(outs) * percent / 100.0))
        picks = rng.sample(outs, min(n, len(outs)))
        picks.sort()
        for op in picks:
            candidates = report_index.get(_match_key(op)) or []
            src = candidates[0] if len(candidates) == 1 else None
            sampled.append((batch, op, src, "c%d" % seq))
            seq += 1

    print("")
    print("=" * 60)
    print("正在生成抽样复核单（每批 %g%%，共 %d 份）..." % (percent, len(sampled)))

    results = {}
    n_bad = 0
    for batch, op, src, _anchor_id in sampled:
        rows, media_rows, err = _compare_pair(src, op)
        results[op] = (rows, media_rows, err)
        mismatched = err or any(not m for *_x, m in rows) or any(not m for *_x, m in media_rows)
        rel = os.path.relpath(op, output_folder)
        if mismatched:
            n_bad += 1
            print("%s\n%s！%s-->检测完毕，发现异常。" % (rel, _RED, _RESET))
        else:
            print("%s\n-->检查完毕，准确无误。" % rel)

    os.makedirs(REVIEW_FOLDER, exist_ok=True)
    html_path = os.path.abspath(os.path.join(
        REVIEW_FOLDER, "复核单_%s.html" % time.strftime("%Y%m%d_%H%M%S")))
    with open(html_path, "w", encoding="utf-8") as fp:
        fp.write(_render_review_html(sampled, results, percent, seed))

    print("")
    print("抽样复核单已生成：%s" % html_path)
    print("请用浏览器/WPS 打开逐份人工复核（红色行为不一致项）。")
    return html_path


def _run(log, tmp_path: str, final_box: List[str]) -> int:
    """返回问题数量（0 表示完全通过）。"""
    report_folder = "检验报告"
    output_folder = "原始记录汇总"

    # 解压 zip（如有）；损坏 zip 计入 fatal_issues，绝不允许假“全部通过”
    extract_root, bad_zips = _ensure_extracted(report_folder)
    fatal_issues = ["压缩包解压失败，源数据不完整，本次检查结果不可信：" + z
                    for z in bad_zips]

    # 递归收集所有原始记录 .docx（跳过临时锁文件；剪枝临时目录）
    output_entries = []
    for root, dirs, files in os.walk(output_folder):
        _prune_ignored_dirs(dirs)
        for f in files:
            if f.lower().endswith('.docx') and not f.startswith('~$'):
                output_entries.append(os.path.join(root, f))
    output_entries.sort()

    total_count = len(output_entries)
    # 每项为 (相对路径, 异常数, 该文件的异常明细行列表)；检查过程中只记不打印
    error_files = []

    print(f"一共检测到有 {total_count} 份原始记录待检查，现在正在进行检查~")
    print("")

    # 查找检验报告的搜索范围（优先解压目录，否则原始报告文件夹）
    search_root = extract_root if extract_root else report_folder

    # 把 .wps/.doc/.docm 源报告转成 .docx（与 main.py 生成时的处理一致；COM 必须在主进程完成）。
    # 始终遍历检验报告根目录：其 walk 会自然覆盖内部解压目录、手动子文件夹与根目录散文件，
    # 保证三种输入法混用时待转换文件无一遗漏。
    _convert_non_docx_reports(report_folder)

    # 批次级双向盘点：源↔成品的批次与文件一一对应（转换完成后再核对）
    # 同时传入 zip 解压目录口径与检验报告根目录口径，确保三种输入法混用时无漏报
    src_batch_maps = [_scan_source_batches(report_folder)]
    if extract_root:
        src_batch_maps.append(_scan_zip_batches(extract_root))
    recon_issues = _reconcile_batches(src_batch_maps, output_folder)
    if recon_issues:
        print(f"！批次双向盘点发现 {len(recon_issues)} 个问题（明细见末尾汇总）：")
        for x in recon_issues[:10]:
            print("  " + x)
        if len(recon_issues) > 10:
            print(f"  ...其余 {len(recon_issues) - 10} 条见末尾")
        print("")

    # 0 份成品必须显式判为异常：空跑“全部通过”是绿屏假象
    if total_count == 0:
        fatal_issues.append("未在“%s”中检测到任何原始记录成品，未执行逐文件比对"
                            % output_folder)

    # ---- 确定最终日志名（同 main.py，前缀 check_；0 份也改名，避免残留 _pending 日志）----
    final_path = os.path.join(
        LOG_FOLDER,
        _build_log_filename([os.path.basename(t) for t in output_entries]))
    final_box[0] = final_path
    sys.stdout = sys.__stdout__
    log.flush()
    log.close()
    try:
        os.replace(tmp_path, final_path)   # 同名直接覆盖：重复检查不产生新日志
    except OSError:
        pass
    log = open(final_path, 'a', encoding='utf-8')
    sys.stdout = _Tee(sys.__stdout__, log)

    # 一次性建立 报告文件名->[路径...] 索引
    # 混用时：zip 解压目录 + 检验报告根目录同时索引；
    # _prune_ignored_dirs 会在遍历 report_folder 时剪枝掉内部解压目录，不会重复。
    index_roots = (search_root, report_folder) if extract_root else (report_folder,)
    report_index = _build_report_index(*index_roots)

    tasks = []
    for output_path in output_entries:
        rel_path = os.path.relpath(output_path, output_folder)
        candidates = report_index.get(_match_key(output_path))
        if not candidates:
            report_path = ""          # 找不到源
        elif len(candidates) == 1:
            report_path = candidates[0]
        else:
            report_path = None        # 同名歧义：fail-closed，不猜
        tasks.append((output_path, rel_path, report_path))

    def _consume(result):
        rel_path, status, count, detail = result
        if status == "ok":
            print(f"正在检查 {rel_path}\n-->检查完毕，准确无误。")
        else:
            print(f"正在检查 {rel_path}\n{_RED}！{_RESET}-->检测完毕，发现异常。")
            error_files.append((rel_path, count, detail))

    if tasks:
        # 纯读取比对可并行（不含 COM）；WPS 转换已在上面主进程完成
        process_count = min(8, multiprocessing.cpu_count())
        print(f"共 {total_count} 份，使用 {process_count} 进程并行检查...\n")
        try:
            with multiprocessing.Pool(processes=process_count) as pool:
                # imap 保序：流式输出顺序与文件排序一致
                for result in pool.imap(_check_one, tasks, chunksize=4):
                    _consume(result)
        # 多进程失败的异常类型不可控（平台/资源相关），回退单线程必须接住
        # noinspection PyBroadException
        except Exception as e:
            # 多进程不可用时回退单线程，保证检查一定能跑完
            print(f"多进程检查失败（{e}），回退单线程执行...\n")
            for task in tasks:
                _consume(_check_one(task))

    # 抽样人工复核：在清理解压目录之前生成（report_index 可能引用解压目录内的文件）
    review_path = None
    try:
        review_path = _generate_sample_review(output_folder, report_index)
    # 抽样复核是附加功能，失败不影响主校验结果
    # noinspection PyBroadException
    except Exception as e:
        print("抽样复核单生成失败（不影响校验结果）: %s" % str(e)[:80])

    # 清理解压临时目录
    if extract_root and os.path.exists(extract_root):
        try:
            shutil.rmtree(extract_root, ignore_errors=True)
            print("\n已清理检查临时解压目录")
        except OSError:
            pass

    print("")
    print("=" * 60)

    if fatal_issues:
        print("！致命问题（本次检查不能判定为通过）：")
        for x in fatal_issues:
            print("  " + x)
        print("")

    if error_files:
        # 末尾统一输出全部异常明细
        print(f"共有 {len(error_files)} 份原始记录数据异常:")
        print("")
        for f, error_count, detail in error_files:
            print(f"{f}（共计有{error_count}个异常数据）-->异常数据如下：")
            for line in detail:
                print(line)
            print("")
    elif not fatal_issues:
        print("逐文件比对：全部通过。")

    # 批次级双向盘点明细（独立于逐文件结果：整批丢失在逐文件视图里不可见）
    print("-" * 60)
    if recon_issues:
        print(f"批次双向盘点：发现 {len(recon_issues)} 个问题，请逐条核查：")
        for x in recon_issues:
            print("  " + x)
    else:
        print("批次双向盘点：源报告与成品批次/数量完全对应，无缺失。")

    # 抽样复核提醒
    if review_path:
        print("-" * 60)
        print("已经生成抽样复核，请人工核对：%s" % review_path)

    return len(error_files) + len(recon_issues) + len(fatal_issues)


def _acquire_single_instance(name: str) -> Optional[int]:
    """命名互斥量：防止两个检查/渲染进程同时运行（互删解压目录、互相污染结果）。

    返回互斥量句柄（进程退出时由 OS 自动释放）；已有实例在运行则返回 None。
    """
    if os.name != 'nt':
        return -1
    try:
        import ctypes
        k32 = ctypes.WinDLL('kernel32', use_last_error=True)
        error_already_exists = 183
        handle = k32.CreateMutexW(None, True, name)
        if not handle:
            return -1  # 创建失败不阻塞业务，仅失去单实例保护
        if ctypes.get_last_error() == error_already_exists:
            return None
        return handle
    except (OSError, AttributeError, TypeError, ValueError):
        return -1


def _enable_ansi_color():
    """在 Windows 传统控制台上启用 VT100 颜色解析（Windows Terminal/PowerShell7 原生支持）。"""
    if os.name != 'nt':
        return
    try:
        import ctypes
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.GetStdHandle(-11)  # STD_OUTPUT_HANDLE
        mode = ctypes.c_uint32()
        if kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
            kernel32.SetConsoleMode(handle, mode.value | 0x0004)  # ENABLE_VIRTUAL_TERMINAL_PROCESSING
    except (OSError, AttributeError):
        pass


def main():
    _enable_ansi_color()
    os.makedirs(LOG_FOLDER, exist_ok=True)
    # 先写临时日志，任务收集齐后再按报告编号范围改名（命名规则同 main.py，前缀 check_）
    tmp_path = os.path.join(LOG_FOLDER, f"_pending_check_{os.getpid()}.log")
    log = open(tmp_path, "w", encoding="utf-8")
    sys.stdout = _Tee(sys.__stdout__, log)
    final_box = [tmp_path]
    problems = 0
    try:
        problems = _run(log, tmp_path, final_box)
    finally:
        sys.stdout = sys.__stdout__
        try:
            log.close()
        except OSError:
            pass
        # 兜底：若中途异常没来得及改名，保留临时文件内容到最终名
        final = final_box[0]
        if final != tmp_path and os.path.exists(tmp_path) and not os.path.exists(final):
            try:
                os.replace(tmp_path, final)
            except OSError:
                pass
    print(f"\n本次检查日志已保存: {final_box[0]}")
    return problems


if __name__ == "__main__":
    _mutex = _acquire_single_instance("Global\\yuanshijilu_check_singleton_v1")
    if _mutex is None:
        print("已有一个 check.py 正在运行，为避免互删解压目录/互相污染结果，本次启动已退出。")
        sys.exit(2)
    sys.exit(1 if main() else 0)
