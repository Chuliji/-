# -*- coding: utf-8 -*-
"""L3 渲染闸门（交付前终极校验）：WPS 排版引擎导出 PDF，附页逐像素比对。

设计依据：附页声称“整页照搬”，那么源报告与成品在同一排版引擎（WPS）下
渲染结果必须逐像素一致。文本/XML 比对无法发现的标点压缩、印章丢失、
样式撞名等亚像素差异，在像素比对中全部现形。

用法：
    python verify.py                  # 附页模式：自动找出全部含附页的成品并比对
    python verify.py 6-8月份结算      # 只检查指定批次（批次名=成品目录名去掉“原始记录”）
    python verify.py --dpi 150        # 指定渲染分辨率（默认 110）
    python verify.py --tol 50         # 允许的差异像素字节数（默认 0，严格）
    python verify.py --full-text      # 全量文本层模式：全部成品的 44 字段+介质列表
                                      # 逐一核对是否出现在成品 PDF 文本层（第三独立通道，
                                      # 不经 python-docx，封死 main/check 共享解析器盲区）
    python verify.py --full-text 6-8月份结算   # 全量文本层模式 + 批次过滤

说明：正文部分模板不同不能整页像素比，由 check.py 的逐字段比对负责；
像素闸门只校验“原样拷贝”的附页，文本层闸门校验全部字段在渲染结果中真实存在。

附页页码按需求已从成品页脚剔除（源报告仍保留），像素比对时以源页页码位置
为准，在双方渲染图上同步遮蔽该小区域后再比；除页码外仍逐像素严格。
"""

# 复用 check 模块的成熟逻辑（同目录；导入时仅执行依赖注册，无其他副作用）
import check as ck
from check import (qn, etree, os, sys, time, shutil, zipfile)
from typing import Dict, List, Optional, Tuple

REPORT_FOLDER = "检验报告"
OUTPUT_FOLDER = "原始记录汇总"
TMP_FOLDER = "_verify_tmp"
LOG_FOLDER = os.path.join("终端结果", "verify运行日志")
HEADING = "充装介质附页"

_ANSI_RE = ck._ANSI_RE
_RED, _RESET = ck._RED, ck._RESET

# 模块级声明：_run 中赋值、_collect_targets 中读取（避免未定义引用告警）
REPORT_INDEX: Dict[str, List[str]] = {}


def _docx_has_appendix(path: str) -> bool:
    """轻量探测：直接读 zip 内 document.xml（不经 python-docx），判断是否含附页。"""
    try:
        with zipfile.ZipFile(path) as z:
            xml_root = etree.fromstring(z.read("word/document.xml"))
    except (zipfile.BadZipFile, KeyError, OSError,
            etree.XMLSyntaxError, ValueError):
        return False
    full = "".join(t.text or "" for t in xml_root.iter(qn("w:t")))
    return HEADING in full


def _collect_targets(batch_filter: Optional[str]) -> List[Tuple[str, Optional[str], str, int]]:
    """返回待比对的 (成品路径, 源路径或None, 相对成品路径) 列表。

    成品含附页才纳入；源路径歧义/缺失时以 None 标记，交后续报错（fail-closed）。
    """
    out_entries = []
    for walk_root, dirs, files in os.walk(OUTPUT_FOLDER):
        for f in files:
            if f.lower().endswith(".docx") and not f.startswith("~$"):
                out_entries.append(os.path.join(walk_root, f))
    out_entries.sort()

    if batch_filter:
        prefix = batch_filter + "原始记录"
        # 按第一级目录名精确匹配，避免“X原始记录备份”被“X”前缀误命中
        out_entries = [p for p in out_entries
                       if os.path.relpath(p, OUTPUT_FOLDER).split(os.sep)[0] == prefix]

    targets = []
    for op in out_entries:
        if not _docx_has_appendix(op):
            continue
        rel = os.path.relpath(op, OUTPUT_FOLDER)
        candidates = REPORT_INDEX.get(ck._match_key(op))
        src = None
        if candidates and len(candidates) == 1:
            src = candidates[0]
        targets.append((op, src, rel, len(candidates) if candidates else 0))
    return targets


def _collect_all_targets(batch_filter: Optional[str]) -> List[Tuple[str, Optional[str], str, int]]:
    """全量文本层模式：纳入所有成品（不筛选附页）。返回结构同 _collect_targets。"""
    out_entries = []
    for walk_root, _dirs, files in os.walk(OUTPUT_FOLDER):
        for f in files:
            if f.lower().endswith(".docx") and not f.startswith("~$"):
                out_entries.append(os.path.join(walk_root, f))
    out_entries.sort()

    if batch_filter:
        prefix = batch_filter + "原始记录"
        # 按第一级目录名精确匹配，避免“X原始记录备份”被“X”前缀误命中
        out_entries = [p for p in out_entries
                       if os.path.relpath(p, OUTPUT_FOLDER).split(os.sep)[0] == prefix]

    targets = []
    for op in out_entries:
        rel = os.path.relpath(op, OUTPUT_FOLDER)
        candidates = REPORT_INDEX.get(ck._match_key(op))
        src = None
        if candidates and len(candidates) == 1:
            src = candidates[0]
        targets.append((op, src, rel, len(candidates) if candidates else 0))
    return targets


def _export_pdfs(word, targets: List[Tuple[str, Optional[str], str, int]],
                 pdf_dir: str, need_src: bool = True) -> Dict[str, str]:
    """用一个 WPS 实例把源报告和成品全部导出为 PDF。返回 {docx路径: pdf路径}。

    need_src=False 时跳过源报告导出（仅全量文本层模式使用：该模式只在成品 PDF
    文本层中查源字段值，源报告由 python-docx 提取，无需渲染，可省一半导出量）。

    逐文件容错：单份损坏/打不开只打印并跳过（后续比对会按"导出失败"落成错误），
    绝不让一份坏文件中断整批渲染闸门。
    """
    mapping = {}
    failed = set()
    seq = 0
    for op, src, rel, nc in targets:
        docs = (src, op) if need_src else (op,)
        for docx in docs:
            if not docx or docx in mapping or docx in failed:
                continue
            seq += 1
            pdf = os.path.join(pdf_dir, "f%03d.pdf" % seq)
            try:
                d = word.Documents.Open(os.path.abspath(docx), ReadOnly=True)
                try:
                    # 17 = wdFormatPDF
                    d.SaveAs2(os.path.abspath(pdf), FileFormat=17)
                finally:
                    d.Close(False)
                mapping[docx] = pdf
            # COM 单文件失败（损坏/加密等）只记录，继续下一份
            # noinspection PyBroadException
            except Exception as e:
                failed.add(docx)
                print(f"  ！PDF 导出失败 {os.path.basename(docx)}: {str(e)[:80]}")
    return mapping


def _appendix_page_range(doc) -> Optional[List[int]]:
    """定位 PDF 中附页页范围：标题文字最后一次命中页起至文末；无则 None。"""
    import pymupdf
    hits = [i for i in range(len(doc)) if doc[i].search_for(HEADING)]
    if not hits:
        return None
    return list(range(hits[-1], len(doc)))


def _count_byte_diff(a: bytes, b: bytes) -> int:
    """统计两段 bytes 的不同字节数（长度不同先报尺寸）。"""
    return sum(1 for x, y in zip(a, b) if x != y)


def _footer_mask_rects(page) -> List[Tuple[float, float, float, float]]:
    """源附页页脚中页码文字的包围矩形（PDF pt，含留白）。

    附页页码按需求只在成品中取消（成品页脚 PAGE 域已被 main 剔除）。
    比对前在双方图像同一位置同步涂白，使闸门除“有意取消的页码”外仍逐像素严格。
    正文底边（pgMar bottom 1440 twips ≈ 0.915h）以下、0.92h 起的文字即页脚区，
    与正文内容天然隔离。
    """
    h = page.rect.height
    rects = []
    for wd in page.get_text("words"):
        x0, y0, x1, y1 = wd[:4]
        if y0 >= h * 0.92:
            rects.append((x0 - 4, y0 - 3, x1 + 4, y1 + 3))
    return rects


def _blank_rects_samples(pix, rects_pt, dpi: int) -> bytes:
    """把 pixmap 中给定矩形（pt 坐标）涂成纯白，返回新 samples bytes。"""
    if not rects_pt:
        return pix.samples
    buf = bytearray(pix.samples)
    scale = dpi / 72.0
    w, h, ncomp = pix.width, pix.height, pix.n
    for x0, y0, x1, y1 in rects_pt:
        px0 = max(0, int(x0 * scale))
        px1 = min(w - 1, int(x1 * scale) + 1)
        py0 = max(0, int(y0 * scale))
        py1 = min(h - 1, int(y1 * scale) + 1)
        for y in range(py0, py1 + 1):
            row = (y * w + px0) * ncomp
            for x in range(px0, px1 + 1):
                i = row + (x - px0) * ncomp
                for c in range(ncomp):
                    buf[i + c] = 255
    return bytes(buf)


def _verify_one(target: Tuple[str, Optional[str], str, int],
                mapping: Dict[str, str], dpi: int, tol: int) -> Tuple[str, List[str]]:
    """比对单份：返回 (status, 明细行)；status ∈ {'ok','error'}"""
    op, src, rel, nc = target
    detail = []

    if not src:
        if nc == 0:
            detail.append("  含附页成品未找到对应的源报告，无法比对")
        else:
            detail.append("  匹配到 %d 份同名源报告（歧义），无法唯一确定" % nc)
        return "error", detail

    pdf_r = mapping.get(src)
    pdf_o = mapping.get(op)
    if not pdf_r or not pdf_o:
        detail.append("  PDF 导出失败（源报告或成品损坏/打不开），无法渲染比对")
        return "error", detail

    import pymupdf
    sr = pymupdf.open(pdf_r)
    so = pymupdf.open(pdf_o)
    try:
        pr = _appendix_page_range(sr)
        po = _appendix_page_range(so)
        if pr is None:
            detail.append("  源报告 PDF 未定位到“充装介质附页”页")
        if po is None:
            detail.append("  成品 PDF 未定位到“充装介质附页”页")
        if pr is None or po is None:
            return "error", detail

        if len(pr) != len(po):
            detail.append("  附页页数不一致：源 %d 页、成品 %d 页" % (len(pr), len(po)))
            return "error", detail

        total_diff = 0
        for a, b in zip(pr, po):
            xa = sr[a].get_pixmap(dpi=dpi)
            xb = so[b].get_pixmap(dpi=dpi)
            if xa.width != xb.width or xa.height != xb.height:
                detail.append("  第 %d 附页渲染尺寸不一致：源 %dx%d、成品 %dx%d"
                              % (a - pr[0] + 1, xa.width, xa.height,
                                 xb.width, xb.height))
                return "error", detail
            # 页码按需求只在成品取消：以源页页码位置为准，双方同步涂白后再比
            masks = _footer_mask_rects(sr[a])
            n = _count_byte_diff(
                _blank_rects_samples(xa, masks, dpi),
                _blank_rects_samples(xb, masks, dpi))
            total_diff += n
            if n:
                detail.append("  第 %d 附页存在 %d 个差异字节（%d dpi，已遮蔽页码区）"
                              % (a - pr[0] + 1, n, dpi))

        if total_diff > tol:
            detail.insert(0, "  附页渲染与源页不一致（差异字节合计 %d，阈值 %d）"
                          % (total_diff, tol))
            return "error", detail
        return "ok", []
    finally:
        sr.close()
        so.close()


def _verify_text_one(target: Tuple[str, Optional[str], str, int],
                     mapping: Dict[str, str]) -> Tuple[str, List[str]]:
    """全量文本层比对单份：源侧 44 字段+介质列表的每个非空值（_norm 归一化后）
    必须出现在成品 PDF 文本层中；缺失即异常。返回 (status, 明细行)。

    独立点：成品一侧不读 docx，而是读 WPS 渲染出的 PDF 文本层——python-docx
    解析盲区（文本框/域代码/隐藏内容等）在渲染结果中原形毕露。
    """
    op, src, rel, nc = target
    detail = []

    if not src:
        if nc == 0:
            detail.append("  成品未找到对应的源报告，无法比对")
        else:
            detail.append("  匹配到 %d 份同名源报告（歧义），无法唯一确定" % nc)
        return "error", detail

    pdf_o = mapping.get(op)
    if not pdf_o:
        detail.append("  成品 PDF 导出失败（损坏/打不开），无法文本比对")
        return "error", detail

    try:
        src_data = ck.extract_report(src)
    # 源报告解析失败 = 无法建立比对基准，按错误处理（fail-closed）
    # noinspection PyBroadException
    except Exception as e:
        detail.append("  源报告字段提取失败: %s" % str(e)[:80])
        return "error", detail

    import pymupdf
    so = pymupdf.open(pdf_o)
    try:
        # 全页文本拼接后去全部空白：单元格内换行/跨页断行均被归一化吸收
        out_text = ck._norm("".join(so[i].get_text() for i in range(len(so))))
    finally:
        so.close()

    missing = []
    for field in ck.FIELDS:
        v = src_data.get(field)
        if not v:
            continue
        nv = ck._norm(str(v))
        if nv and nv not in out_text:
            missing.append("    %s = %s" % (field, v))
    for row in src_data.get("适装介质列表", []):
        name = ck._norm(row.get("介质名称", ""))
        if name and name not in out_text:
            missing.append("    适装介质[%s] = %s"
                           % (row.get("序号", "?"), row.get("介质名称", "")))

    if missing:
        detail.append("  以下源字段值未在成品 PDF 文本层出现"
                      "（如肉眼复核 PDF 实际存在，属换行排版误报，可忽略）：")
        detail.extend(missing)
        return "error", detail
    return "ok", []


def _parse_args(argv: List[str]) -> Tuple[Optional[str], int, int, bool]:
    batch, dpi, tol, full_text = None, 110, 0, False
    i = 0
    while i < len(argv):
        a = argv[i]
        if a == "--dpi":
            dpi = int(argv[i + 1]); i += 2; continue
        if a == "--tol":
            tol = int(argv[i + 1]); i += 2; continue
        if a == "--full-text":
            full_text = True; i += 1; continue
        batch = a
        i += 1
    return batch, dpi, tol, full_text


def _run(log, tmp_path: str, final_box: List[str]) -> int:
    global REPORT_INDEX

    batch_filter, dpi, tol, full_text = _parse_args(sys.argv[1:])

    # 启动前清理上次残留的临时 PDF（防磁盘占用、防止旧文件干扰当次结果）
    if os.path.exists(TMP_FOLDER):
        shutil.rmtree(TMP_FOLDER, ignore_errors=True)

    # 源报告准备：解压 zip + 非 docx 转换（复用 check 逻辑）。
    # 使用独立解压标签 _verify，避免与 check 的 _已解压_check 冲突（防止同时跑时互删）
    extract_root, bad_zips = ck._ensure_extracted(REPORT_FOLDER, extract_tag="_已解压_verify")
    search_root = extract_root if extract_root else REPORT_FOLDER
    # 始终遍历检验报告根目录做转换，确保三种输入法混用时不遗漏 .wps 文件
    ck._convert_non_docx_reports(REPORT_FOLDER)
    # 同时索引两种根：zip 解压目录 + 检验报告根目录（混用时无漏报）；
    # _prune_ignored_dirs 会剪枝掉内部解压目录，不会重复
    index_roots = (search_root, REPORT_FOLDER) if extract_root else (REPORT_FOLDER,)
    REPORT_INDEX = ck._build_report_index(*index_roots)

    if full_text:
        targets = _collect_all_targets(batch_filter)
        mode_label = "全量文本层比对"
    else:
        targets = _collect_targets(batch_filter)
        mode_label = "附页渲染闸门"

    # 日志命名：verify_ + 编号范围（同 check 口径）；0 份也改名，避免残留 _pending 日志
    log_prefix = "verify_fulltext_" if full_text else "verify_"
    final_name = ck._build_log_filename(
        [os.path.basename(t[0]) for t in targets]).replace("check_", log_prefix, 1)
    final_path = os.path.join(LOG_FOLDER, final_name)
    final_box[0] = final_path
    sys.stdout = sys.__stdout__
    log.flush(); log.close()
    try:
        os.replace(tmp_path, final_path)
    except OSError:
        pass
    log = open(final_path, "a", encoding="utf-8")
    sys.stdout = ck._Tee(sys.__stdout__, log)

    if bad_zips:
        print("！压缩包解压失败，源数据不完整，本次%s结果不可信：" % mode_label)
        for z in bad_zips:
            print("  " + z)
    if not targets:
        if full_text:
            print("未找到成品，无需比对。")
        else:
            print("未找到含充装介质附页的成品，无需比对。")
        if extract_root and os.path.exists(extract_root):
            shutil.rmtree(extract_root, ignore_errors=True)
        return len(bad_zips)

    if full_text:
        print(f"共 {len(targets)} 份成品，逐字段核对源值是否出现在成品 PDF 文本层，"
              f"正在导出 PDF（数量较多，耗时较长）...")
    else:
        print(f"共 {len(targets)} 份含附页成品，按 {dpi} dpi 渲染比对"
              f"（差异阈值 {tol}），正在导出 PDF...")

    pdf_dir = os.path.join(TMP_FOLDER, "pdf")
    os.makedirs(pdf_dir, exist_ok=True)
    errors = []

    import win32com.client
    word = None
    try:
        word = win32com.client.Dispatch("KWps.Application")
        word.Visible = False
        try:
            word.DisplayAlerts = 0
        # COM 属性异常不影响导出
        # noinspection PyBroadException
        except Exception:
            pass
        mapping = _export_pdfs(word, targets, pdf_dir, need_src=not full_text)
        word.Quit()
        word = None

        if full_text:
            print(f"PDF 导出完成（{len(mapping)} 个文档），开始文本层比对...\n")
        else:
            print(f"PDF 导出完成（{len(mapping)} 个文档），开始逐页像素比对...\n")
        for t in targets:
            if full_text:
                status, detail = _verify_text_one(t, mapping)
            else:
                status, detail = _verify_one(t, mapping, dpi, tol)
            rel = t[2]
            if status == "ok":
                if full_text:
                    print(f"正在校验 {rel}\n-->检查完毕，准确无误。")
                else:
                    print(f"正在校验 {rel}\n-->渲染一致，附页与源页逐像素相同（页码区按需求已遮蔽）。")
            else:
                print(f"正在校验 {rel}\n{_RED}！{_RESET}-->{mode_label}发现差异。")
                errors.append((rel, detail))
    finally:
        if word is not None:
            try:
                word.Quit()
            # noinspection PyBroadException
            except Exception:
                pass

    print("")
    print("=" * 60)
    if errors:
        print(f"{mode_label}：{len(errors)} 份发现差异，交付前必须处理：")
        print("")
        for rel, detail in errors:
            print(rel + "：")
            for line in detail:
                print(line)
            print("")
    elif not bad_zips:
        if full_text:
            print(f"{mode_label}：全部 {len(targets)} 份成品的源字段值均在 PDF 文本层出现，通过。")
        else:
            print(f"{mode_label}：全部 {len(targets)} 份与源报告逐像素一致，通过。")

    # 清理全部临时文件（独立解压目录 _已解压_verify + PDF 目录）
    for d in (extract_root, TMP_FOLDER):
        if d and os.path.exists(d):
            shutil.rmtree(d, ignore_errors=True)
    print("已清理临时目录。")

    return len(errors) + len(bad_zips)


def main():
    _mutex = ck._acquire_single_instance("Global\\yuanshijilu_verify_singleton_v1")
    if _mutex is None:
        print("已有一个 verify.py 正在运行，为避免互删解压目录/互相污染结果，本次启动已退出。")
        return 2
    ck._enable_ansi_color()
    os.makedirs(LOG_FOLDER, exist_ok=True)
    tmp_path = os.path.join(LOG_FOLDER, f"_pending_verify_{os.getpid()}.log")
    log = open(tmp_path, "w", encoding="utf-8")
    sys.stdout = ck._Tee(sys.__stdout__, log)
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
        final = final_box[0]
        if final != tmp_path and os.path.exists(tmp_path) and not os.path.exists(final):
            try:
                os.replace(tmp_path, final)
            except OSError:
                pass
    print(f"\n本次渲染闸门日志已保存: {final_box[0]}")
    return problems


if __name__ == "__main__":
    sys.exit(1 if main() else 0)
