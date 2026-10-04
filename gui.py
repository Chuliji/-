# -*- coding: utf-8 -*-
"""原始记录自动生成（顺德）—— 桌面 GUI 外壳。

本文件只做"外壳 GUI + 脚本调用"，不改动 main.py / check.py / verify.py /
run_all.py 里的任何业务逻辑：
  - 四个业务脚本原样通过子进程调用（frozen 时调用本 exe，由本入口转 runpy 执行）；
  - 模板 原始记录模版2.0.docx 与 _deps/ 随 PyInstaller 内置，脚本自身已有
    sys._MEIPASS 处理，直接沿用；
  - 业务目录约定不变：检验报告/ 原始记录汇总/ 抽样复核/ 终端结果/。

特殊进程角色（frozen 时由 PyInstaller bootloader 拉起）：
  1. app.exe --multiprocessing-fork ...   main.py 进程池的 worker 子进程；
     需先按环境变量 _MP_RUNNER_SCRIPT 预载脚本定义，再 multiprocessing.freeze_support()。
  2. app.exe <脚本名>.py [args...]        运行业务脚本（run_all.py 内部就是这样
     用 sys.executable 调 check.py / verify.py 的）。
  3. 无参数                                启动 GUI。
"""

# ============================================================================
# 第 0 步：自动注册本地 _deps 依赖（与业务脚本同一段引导逻辑，pywin32 等）
# ============================================================================
import os as _os
import sys as _sys


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
        except Exception:
            pass

# 业务脚本运行时需要的三方库：在此显式导入，保证 PyInstaller 分析时能收集到
# （脚本本身是 --add-data 数据文件，不会被 PyInstaller 静态分析）
import docx  # noqa: E402,F401
import lxml.etree  # noqa: E402,F401
import fitz  # noqa: E402,F401  (PyMuPDF)

import runpy  # noqa: E402
import shutil  # noqa: E402
import webbrowser  # noqa: E402
from pathlib import Path  # noqa: E402


# ============================================================================
# 路径与运行时脚本准备
# ============================================================================
def _base_dir() -> Path:
    """业务工作目录（放 检验报告/ 等）：frozen 时在 exe 同级，开发时在项目根。"""
    if getattr(_sys, "frozen", False):
        return Path(_sys.executable).resolve().parent
    return Path(__file__).resolve().parent


BASE_DIR = _base_dir()

SCRIPTS = ("main.py", "check.py", "verify.py", "run_all.py")
BUSINESS_DIRS = ("检验报告", "原始记录汇总", "抽样复核", "终端结果")


def _resource_dir() -> Path:
    """PyInstaller 内置资源根（_MEIPASS）；开发时即项目根。"""
    if hasattr(_sys, "_MEIPASS"):
        return Path(_sys._MEIPASS)
    return BASE_DIR


def ensure_dirs() -> None:
    for d in BUSINESS_DIRS:
        (BASE_DIR / d).mkdir(parents=True, exist_ok=True)


def ensure_runtime_scripts() -> None:
    """frozen 时把业务脚本从 _MEIPASS 同步到 BASE_DIR。

    run_all.py 内部用 cwd=dirname(__file__) + sys.executable 调子脚本，
    所以脚本必须位于 BASE_DIR（与开发时的项目布局完全一致），
    这样 run_all 的 __file__ 推导、cwd、脚本互相 import 都和现在一模一样。
    """
    if not getattr(_sys, "frozen", False):
        return
    res = _resource_dir()
    for name in SCRIPTS:
        src = res / name
        if not src.exists():
            continue
        dst = BASE_DIR / name
        if dst.exists():
            # 内容一致就不重写，避免无谓的磁盘/杀软干扰
            try:
                if src.read_bytes() == dst.read_bytes():
                    continue
            except OSError:
                pass
        try:
            shutil.copy2(src, dst)
        except OSError:
            pass


def _resolve_script(script_name: str) -> str:
    """按 cwd → BASE_DIR → _MEIPASS 顺序找到业务脚本。"""
    candidates = [
        Path.cwd() / script_name,
        BASE_DIR / script_name,
        _resource_dir() / script_name,
    ]
    for p in candidates:
        if p.is_file():
            return str(p)
    # 给绝对路径一个机会
    p = Path(script_name)
    if p.is_file():
        return str(p.resolve())
    raise FileNotFoundError(f"找不到业务脚本: {script_name}")


# ============================================================================
# frozen 子进程：main.py 进程池 worker（必须在导入 GUI 之前拦截）
# ============================================================================
def _mp_worker_bootstrap() -> None:
    if not (len(_sys.argv) >= 2 and _sys.argv[1] == "--multiprocessing-fork"):
        return
    script = _os.environ.get("_MP_RUNNER_SCRIPT")
    if script and _os.path.isfile(script):
        try:
            ns = runpy.run_path(script, run_name="__mp_worker_preload__")
            main_mod = _sys.modules["__main__"]
            for k, v in ns.items():
                if not k.startswith("__"):
                    main_mod.__dict__[k] = v
        except Exception:
            pass
    import multiprocessing
    multiprocessing.freeze_support()
    _sys.exit(0)


# ============================================================================
# frozen 子进程：业务脚本运行器（app.exe <script>.py [args...]）
# ============================================================================
def _run_script(script_name: str, script_args: list) -> None:
    resolved = _resolve_script(script_name)
    # 告知将来 spawn 出的进程池子进程该预载哪个脚本
    _os.environ["_MP_RUNNER_SCRIPT"] = resolved
    _sys.argv = [resolved] + list(script_args)
    try:
        runpy.run_path(resolved, run_name="__main__")
    except SystemExit as e:
        code = e.code
        if code is None:
            code = 0
        elif isinstance(code, str):
            print(code)
            code = 1
        _sys.exit(code)


if len(_sys.argv) >= 2 and _sys.argv[1] == "--multiprocessing-fork":
    _mp_worker_bootstrap()

if len(_sys.argv) >= 2 and _sys.argv[1].endswith(".py"):
    _run_script(_sys.argv[1], _sys.argv[2:])


# ============================================================================
# 以下仅 GUI 主进程使用
# ============================================================================
from PySide6.QtCore import (Qt, QThread, Signal, QProcess, QTimer)
from PySide6.QtGui import (QFont, QTextCursor, QTextCharFormat, QColor,
                           QAction, QGuiApplication, QIcon)
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget, QVBoxLayout, QHBoxLayout,
    QGridLayout, QPushButton, QLabel, QCheckBox, QLineEdit, QPlainTextEdit,
    QFrame, QListWidget, QListWidgetItem, QStatusBar, QProgressBar,
    QMessageBox, QSplitter, QSizePolicy,
)


ACCEPTED_DOC_EXTS = (".docx", ".doc", ".wps", ".docm")


# ---------------------------------------------------------------------------
# 字节流解码：UTF-8 / GBK 自适应（管道输出编码两种都可能遇到）
# ---------------------------------------------------------------------------
class PipeDecoder:
    def __init__(self):
        self._pending = b""

    @staticmethod
    def _incomplete_utf8_tail(data: bytes):
        """返回末尾不完整 UTF-8 序列的起始下标；没有则 None。"""
        i = len(data)
        while i > 0 and 0x80 <= data[i - 1] <= 0xBF:
            i -= 1
        if i == 0:
            # 整段都是续字节，按不完整处理（极端情况）
            return 0
        lead = data[i - 1]
        if lead < 0xC0:
            return None
        need = 1
        if lead >= 0xE0:
            need += 1
        if lead >= 0xF0:
            need += 1
        if len(data) - i < need:
            return i - 1
        return None

    def feed(self, data: bytes) -> str:
        data = self._pending + data
        self._pending = b""
        try:
            return data.decode("utf-8")
        except UnicodeDecodeError:
            cut = self._incomplete_utf8_tail(data)
            if cut is not None and cut > 0:
                self._pending = data[cut:]
                return data[:cut].decode("utf-8", errors="replace")
            return data.decode("gbk", errors="replace")

    def flush(self) -> str:
        if not self._pending:
            return ""
        text = self._pending.decode("gbk", errors="replace")
        self._pending = b""
        return text


# ---------------------------------------------------------------------------
# ANSI 彩色终端
# ---------------------------------------------------------------------------
_ANSI_FG = {
    30: "#8a8f99", 31: "#ff6b6b", 32: "#6ce38a", 33: "#ffd166",
    34: "#6aa9ff", 35: "#c792ea", 36: "#62e6d4", 37: "#e6e6e6",
    90: "#9aa0aa", 91: "#ff8585", 92: "#88ec9f", 93: "#ffdc7a",
    94: "#8db9ff", 95: "#d8a9f0", 96: "#82eee0", 97: "#ffffff",
}


class AnsiTerminal(QPlainTextEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setReadOnly(True)
        self.setMaximumBlockCount(50000)
        self.setFrameStyle(QFrame.NoFrame)
        term_font = QFont("Consolas", 11)
        term_font.setStyleHint(QFont.Monospace)
        term_font.setPointSize(14)
        self.setFont(term_font)
        self._fmt = QTextCharFormat()
        self._fmt.setForeground(QColor("#dfe3ea"))
        self.document().setDocumentMargin(10)

    def clearTerminal(self):
        self.clear()

    def _apply_sgr(self, param: str):
        codes = param.split(";") if param else ["0"]
        for token in codes:
            try:
                c = int(token)
            except ValueError:
                continue
            if c == 0:
                self._fmt = QTextCharFormat()
                self._fmt.setForeground(QColor("#dfe3ea"))
            elif c == 1:
                self._fmt.setFontWeight(QFont.Bold)
            elif c == 22:
                self._fmt.setFontWeight(QFont.Normal)
            elif c in _ANSI_FG:
                self._fmt.setForeground(QColor(_ANSI_FG[c]))
            elif c == 39:
                self._fmt.setForeground(QColor("#dfe3ea"))
            # 背景码（40-49）等本工具脚本未使用，忽略

    def append_ansi(self, text: str):
        if not text:
            return
        sb = self.verticalScrollBar()
        at_bottom = sb.value() >= sb.maximum() - 2
        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.End)
        begin = 0
        import re
        for m in re.finditer(r"\x1b\[([0-9;]*)m", text):
            if m.start() > begin:
                cursor.insertText(text[begin:m.start()], self._fmt)
            self._apply_sgr(m.group(1))
            begin = m.end()
        if begin < len(text):
            cursor.insertText(text[begin:], self._fmt)
        if at_bottom:
            sb.setValue(sb.maximum())


# ---------------------------------------------------------------------------
# WPS COM 检测线程
# ---------------------------------------------------------------------------
class WpsCheckThread(QThread):
    result = Signal(bool)

    def run(self):
        try:
            import win32com.client
            app = win32com.client.Dispatch("KWps.Application")
            try:
                app.Visible = False
            except Exception:
                pass
            try:
                app.Quit()
            except Exception:
                pass
            self.result.emit(True)
        except Exception:
            self.result.emit(False)


# ---------------------------------------------------------------------------
# 拖放区（视觉提示；窗口级事件统一处理）
# ---------------------------------------------------------------------------
class DropZone(QFrame):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("dropZone")
        self.setFixedHeight(132)
        self.setAcceptDrops(True)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 18, 20, 18)
        title = QLabel("将 zip / Word 文档 / 文件夹 拖到这里")
        title.setAlignment(Qt.AlignCenter)
        title.setObjectName("dropTitle")
        sub = QLabel("支持一次拖入多个文件 + 多个文件夹  ·  .zip .docx .doc .wps .docm")
        sub.setAlignment(Qt.AlignCenter)
        sub.setObjectName("dropSub")
        lay.addWidget(title)
        lay.addWidget(sub)
        self._title = title
        self._sub = sub

    def set_active(self, active: bool):
        self.setProperty("active", "1" if active else "")
        self.style().unpolish(self)
        self.style().polish(self)

    def show_intake(self, summary: str):
        self._sub.setText(summary)


# ---------------------------------------------------------------------------
# 主窗口
# ---------------------------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("原始记录自动生成（顺德）· 桌面版")
        self.resize(1180, 880)
        self.setMinimumSize(1000, 720)
        self.setAcceptDrops(True)

        self.process = QProcess(self)
        self.process.setWorkingDirectory(str(BASE_DIR))
        self.process.setProcessChannelMode(QProcess.MergedChannels)
        self.process.readyReadStandardOutput.connect(self._on_process_output)
        self.process.finished.connect(self._on_process_finished)
        self.process.errorOccurred.connect(self._on_process_error)
        self.decoder = PipeDecoder()

        self.wps_thread = WpsCheckThread()
        self.wps_thread.result.connect(self._on_wps_checked)

        self._running_label = ""
        self._wps_ok: bool | None = None
        self._all_controls: list = []

        self._build_ui()
        self.refresh_products()
        self.statusBar().showMessage("就绪")
        # 启动后后台检测 WPS
        QTimer.singleShot(300, self.wps_thread.start)

    # ---------------- UI 构建 ----------------
    def _build_ui(self):
        tabs = QTabWidget()
        tabs.addTab(self._build_run_tab(), "运行台")
        tabs.addTab(self._build_result_tab(), "结果浏览")
        self.setCentralWidget(tabs)
        self.tabs = tabs

        bar = QStatusBar()
        self.setStatusBar(bar)

        self.busy_bar = QProgressBar()
        self.busy_bar.setRange(0, 0)
        self.busy_bar.setFixedWidth(170)
        self.busy_bar.setTextVisible(False)
        self.busy_bar.hide()

        self.run_state_label = QLabel("")
        self.exit_code_label = QLabel("尚未运行")
        self.wps_label = QLabel("WPS：检测中…")
        self.wps_label.setObjectName("wpsChecking")

        bar.addWidget(self.run_state_label, 1)
        bar.addWidget(self.exit_code_label)
        bar.addPermanentWidget(self.busy_bar)
        bar.addPermanentWidget(self.wps_label)

    def _build_run_tab(self) -> QWidget:
        page = QWidget()
        root = QVBoxLayout(page)
        root.setContentsMargins(20, 16, 20, 14)
        root.setSpacing(12)

        # ---- 大按钮区 ----
        grid = QGridLayout()
        grid.setSpacing(12)
        self.btn_main = self._big_button("生成原始记录", "btnBlue")
        self.btn_check = self._big_button("数据校验 check", "btnTeal")
        self.btn_verify = self._big_button("渲染闸门 verify", "btnViolet")
        self.btn_auto = self._big_button("一键全自动", "btnAmber")
        self.btn_main.clicked.connect(lambda: self.start_run("main.py", [], "生成原始记录"))
        self.btn_check.clicked.connect(lambda: self.start_run("check.py", [], "数据校验 check"))
        self.btn_verify.clicked.connect(self._start_verify)
        self.btn_auto.clicked.connect(self._start_auto)
        grid.addWidget(self.btn_main, 0, 0)
        grid.addWidget(self.btn_check, 0, 1)
        grid.addWidget(self.btn_verify, 1, 0)
        grid.addWidget(self.btn_auto, 1, 1)
        root.addLayout(grid)

        # ---- verify 选项行 ----
        v_card = QFrame()
        v_card.setObjectName("optionCard")
        v = QHBoxLayout(v_card)
        v.setContentsMargins(16, 10, 16, 10)
        v_lbl = QLabel("verify 选项：")
        self.chk_fulltext = QCheckBox("全量文本层(--full-text)")
        v.addWidget(v_lbl)
        v.addWidget(self.chk_fulltext)
        v.addWidget(self._small_label("批次名"))
        self.edit_batch = QLineEdit()
        self.edit_batch.setPlaceholderText("可空，透传给 verify")
        self.edit_batch.setFixedWidth(200)
        v.addWidget(self.edit_batch)
        v.addWidget(self._small_label("DPI"))
        self.edit_dpi = QLineEdit("110")
        self.edit_dpi.setFixedWidth(70)
        v.addWidget(self.edit_dpi)
        v.addWidget(self._small_label("容差"))
        self.edit_tol = QLineEdit("0")
        self.edit_tol.setFixedWidth(70)
        v.addWidget(self.edit_tol)
        v.addStretch(1)
        root.addWidget(v_card)

        # ---- 一键全自动选项行 ----
        a_card = QFrame()
        a_card.setObjectName("optionCard")
        a = QHBoxLayout(a_card)
        a.setContentsMargins(16, 10, 16, 10)
        a_lbl = QLabel("全自动选项：")
        self.chk_auto_fulltext = QCheckBox("--full-text")
        self.chk_skip_check = QCheckBox("跳过 check")
        self.chk_skip_verify = QCheckBox("跳过 verify")
        a.addWidget(a_lbl)
        a.addWidget(self.chk_auto_fulltext)
        a.addWidget(self.chk_skip_check)
        a.addWidget(self.chk_skip_verify)
        a.addWidget(self._small_label("批次名透传"))
        self.edit_auto_batch = QLineEdit()
        self.edit_auto_batch.setPlaceholderText("可空")
        self.edit_auto_batch.setFixedWidth(200)
        a.addWidget(self.edit_auto_batch)
        a.addStretch(1)
        root.addWidget(a_card)

        # ---- 拖放区 ----
        self.drop_zone = DropZone()
        root.addWidget(self.drop_zone)

        # ---- 终端标题行 ----
        term_head = QHBoxLayout()
        term_title = QLabel("实时终端")
        term_title.setObjectName("sectionTitle")
        term_head.addWidget(term_title)
        term_head.addStretch(1)
        btn_clear = QPushButton("清空终端")
        btn_clear.setObjectName("ghostButton")
        btn_clear.setCursor(Qt.PointingHandCursor)
        btn_clear.clicked.connect(lambda: self.terminal.clearTerminal())
        term_head.addWidget(btn_clear)
        root.addLayout(term_head)

        self.terminal = AnsiTerminal()
        root.addWidget(self.terminal, 1)

        self._all_controls = [
            self.btn_main, self.btn_check, self.btn_verify, self.btn_auto,
            self.chk_fulltext, self.edit_batch, self.edit_dpi, self.edit_tol,
            self.chk_auto_fulltext, self.chk_skip_check, self.chk_skip_verify,
            self.edit_auto_batch,
        ]
        return page

    def _build_result_tab(self) -> QWidget:
        page = QWidget()
        root = QVBoxLayout(page)
        root.setContentsMargins(20, 16, 20, 14)
        root.setSpacing(12)

        row = QHBoxLayout()
        row.setSpacing(10)
        self.btn_open_out = QPushButton("打开成品目录")
        self.btn_open_html = QPushButton("打开抽样复核 HTML")
        self.btn_open_log = QPushButton("打开日志目录")
        for b in (self.btn_open_out, self.btn_open_html, self.btn_open_log):
            b.setObjectName("resultButton")
            b.setCursor(Qt.PointingHandCursor)
            b.setMinimumHeight(52)
            row.addWidget(b, 1)
        self.btn_open_out.clicked.connect(lambda: self._open_dir(BASE_DIR / "原始记录汇总"))
        self.btn_open_html.clicked.connect(self.open_latest_review_html)
        self.btn_open_log.clicked.connect(lambda: self._open_dir(BASE_DIR / "终端结果"))
        root.addLayout(row)

        head = QHBoxLayout()
        lbl = QLabel("成品列表（双击某条 → 复制完整路径到剪贴板）")
        lbl.setObjectName("sectionTitle")
        head.addWidget(lbl)
        head.addStretch(1)
        btn_refresh = QPushButton("刷新列表")
        btn_refresh.setObjectName("ghostButton")
        btn_refresh.setCursor(Qt.PointingHandCursor)
        btn_refresh.clicked.connect(self.refresh_products)
        head.addWidget(btn_refresh)
        root.addLayout(head)

        self.product_list = QListWidget()
        self.product_list.itemDoubleClicked.connect(self._copy_product_path)
        root.addWidget(self.product_list, 1)

        self.product_count_label = QLabel("")
        root.addWidget(self.product_count_label)
        return page

    def _big_button(self, text: str, obj_name: str) -> QPushButton:
        btn = QPushButton(text)
        btn.setObjectName(obj_name)
        btn.setMinimumHeight(68)
        btn.setCursor(Qt.PointingHandCursor)
        btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        return btn

    @staticmethod
    def _small_label(text: str) -> QLabel:
        lbl = QLabel(text)
        lbl.setObjectName("smallLabel")
        return lbl

    # ---------------- 运行控制 ----------------
    def _set_busy(self, busy: bool):
        for w in self._all_controls:
            w.setDisabled(busy)
        self.busy_bar.setVisible(busy)
        self.drop_zone.setEnabled(not busy)

    def start_run(self, script: str, args: list, label: str):
        if self.process.state() != QProcess.NotRunning:
            QMessageBox.information(self, "提示", "已有任务正在运行，请等待结束。")
            return
        self._running_label = label
        cmd = [script] + list(args)
        self._set_busy(True)
        self.tabs.setCurrentIndex(0)
        self.run_state_label.setText(f"正在运行：{label}")
        self.exit_code_label.setText("运行中…")
        self.exit_code_label.setObjectName("")
        self.exit_code_label.style().unpolish(self.exit_code_label)
        self.exit_code_label.style().polish(self.exit_code_label)
        self.terminal.append_ansi(
            f"\n\x1b[1;94m{'=' * 60}\n  ▶ 开始: {label}\n  命令: {' '.join(cmd)}\n{'=' * 60}\x1b[0m\n")
        self.decoder = PipeDecoder()
        # dev: [python.exe, script]；frozen: [app.exe, script]（入口识别 .py 转 runpy）
        self.process.start(_sys.executable, cmd)

    def _start_verify(self):
        args = []
        if self.chk_fulltext.isChecked():
            args.append("--full-text")
        try:
            dpi = int(self.edit_dpi.text().strip() or "110")
            tol = int(self.edit_tol.text().strip() or "0")
        except ValueError:
            QMessageBox.warning(self, "参数有误", "DPI 和容差必须是整数。")
            return
        args += ["--dpi", str(dpi), "--tol", str(tol)]
        batch = self.edit_batch.text().strip()
        if batch:
            args.append(batch)
        self.start_run("verify.py", args, "渲染闸门 verify")

    def _start_auto(self):
        args = []
        if self.chk_auto_fulltext.isChecked():
            args.append("--full-text")
        if self.chk_skip_check.isChecked():
            args.append("--skip-check")
        if self.chk_skip_verify.isChecked():
            args.append("--skip-verify")
        batch = self.edit_auto_batch.text().strip()
        if batch:
            args.append(batch)
        self.start_run("run_all.py", args, "一键全自动")

    def _on_process_output(self):
        data = bytes(self.process.readAllStandardOutput())
        self.terminal.append_ansi(self.decoder.feed(data))

    def _on_process_error(self, _err):
        # start 失败等场景
        self.terminal.append_ansi(f"\x1b[91m启动失败: {self.process.errorString()}\x1b[0m\n")

    def _on_process_finished(self, code: int, status):
        tail = self.decoder.flush()
        if tail:
            self.terminal.append_ansi(tail)
        self._set_busy(False)
        crashed = status == QProcess.CrashExit
        if crashed:
            self.exit_code_label.setText(f"进程异常终止（退出码 {code}）")
            self.exit_code_label.setObjectName("exitFail")
            self.terminal.append_ansi(
                f"\x1b[1;91m✗ {self._running_label} 异常终止，退出码 {code}\x1b[0m\n")
        elif code == 0:
            self.exit_code_label.setText("✓ 全部通过（退出码 0）")
            self.exit_code_label.setObjectName("exitOk")
            self.terminal.append_ansi(
                f"\x1b[1;32m✓ {self._running_label} 全部通过\x1b[0m\n")
        else:
            self.exit_code_label.setText(f"✗ 失败 退出码={code}")
            self.exit_code_label.setObjectName("exitFail")
            self.terminal.append_ansi(
                f"\x1b[1;91m✗ {self._running_label} 失败，退出码 {code}\x1b[0m\n")
        self.exit_code_label.style().unpolish(self.exit_code_label)
        self.exit_code_label.style().polish(self.exit_code_label)
        self.run_state_label.setText(f"上次运行：{self._running_label}")
        self.refresh_products()

    # ---------------- WPS 检测 ----------------
    def _on_wps_checked(self, ok: bool):
        self._wps_ok = ok
        if ok:
            self.wps_label.setText("WPS：✓ 已检测到")
            self.wps_label.setObjectName("wpsOk")
        else:
            self.wps_label.setText("WPS：✗ 未安装")
            self.wps_label.setObjectName("wpsFail")
        self.wps_label.style().unpolish(self.wps_label)
        self.wps_label.style().polish(self.wps_label)
        if not ok:
            QMessageBox.warning(
                self,
                "未检测到 WPS Office",
                "本机未检测到 WPS Office（KWps.Application 组件不可用）。\n\n"
                "本程序的 .doc/.wps 转换、附页页码计算、PDF 导出均依赖 WPS，\n"
                "请先安装 WPS Office 后再使用。\n\n"
                "（安装完成后重新启动本程序即可，无需额外配置。）")

    # ---------------- 拖放归类 ----------------
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.drop_zone.set_active(True)
        else:
            event.ignore()

    def dragLeaveEvent(self, event):
        self.drop_zone.set_active(False)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        self.drop_zone.set_active(False)
        urls = event.mimeData().urls()
        paths = [u.toLocalFile() for u in urls if u.isLocalFile()]
        if not paths:
            return
        event.acceptProposedAction()
        self.ingest_paths(paths)

    def ingest_paths(self, paths: list):
        reports_root = BASE_DIR / "检验报告"
        reports_root.mkdir(parents=True, exist_ok=True)

        n_zip = n_folder = n_doc = n_other = 0
        errors = []

        for raw in paths:
            src = Path(raw)
            try:
                if src.is_dir():
                    dst = reports_root / src.name
                    # 同名文件夹按同批次合并
                    shutil.copytree(src, dst, dirs_exist_ok=True)
                    n_folder += 1
                elif src.is_file():
                    ext = src.suffix.lower()
                    if ext == ".zip":
                        self._copy_file_unique(src, reports_root)
                        n_zip += 1
                    elif ext in ACCEPTED_DOC_EXTS:
                        self._copy_file_unique(src, reports_root)
                        n_doc += 1
                    else:
                        n_other += 1
            except Exception as e:
                errors.append(f"{src.name}: {str(e)[:80]}")

        parts = []
        if n_zip:
            parts.append(f"{n_zip} 个 zip")
        if n_doc:
            parts.append(f"{n_doc} 个文档")
        if n_folder:
            parts.append(f"{n_folder} 个批次文件夹")
        summary = "已识别 " + "、".join(parts) if parts else "没有可处理的输入"
        if n_other:
            summary += f"（{n_other} 个不支持的文件已忽略）"
        self.statusBar().showMessage(summary, 20000)
        self.drop_zone.show_intake(summary)

        if errors:
            QMessageBox.warning(self, "部分项目复制失败",
                                summary + "\n\n失败详情：\n" + "\n".join(errors[:12]))

    @staticmethod
    def _copy_file_unique(src: Path, dst_dir: Path):
        dst = dst_dir / src.name
        if not dst.exists():
            shutil.copy2(src, dst)
            return
        # 同名文件：加序号，避免覆盖
        stem, ext = src.stem, src.suffix
        i = 1
        while True:
            cand = dst_dir / f"{stem}({i}){ext}"
            if not cand.exists():
                shutil.copy2(src, cand)
                return
            i += 1

    # ---------------- 结果浏览 ----------------
    def refresh_products(self):
        if not hasattr(self, "product_list"):
            return
        self.product_list.clear()
        out_root = BASE_DIR / "原始记录汇总"
        files = []
        if out_root.exists():
            for p in out_root.rglob("*.docx"):
                if not p.name.startswith("~$"):
                    files.append(p)
        files.sort(key=lambda p: str(p).lower())
        for p in files:
            item = QListWidgetItem(str(p))
            item.setToolTip(str(p))
            self.product_list.addItem(item)
        self.product_count_label.setText(f"共 {len(files)} 个成品 docx")

    def _copy_product_path(self, item: QListWidgetItem):
        path = item.text()
        QGuiApplication.clipboard().setText(path)
        self.statusBar().showMessage(f"已复制到剪贴板：{path}", 12000)

    @staticmethod
    def _open_dir(path: Path):
        path.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(str(path))  # type: ignore[attr-defined]
        except Exception as e:
            QMessageBox.warning(None, "打开失败", str(e))

    def open_latest_review_html(self):
        review_root = BASE_DIR / "抽样复核"
        if not review_root.exists():
            QMessageBox.information(self, "暂无复核单",
                                    "抽样复核目录不存在，请先运行【数据校验 check】。")
            return
        htmls = [p for p in review_root.rglob("*.html")]
        if not htmls:
            QMessageBox.information(self, "暂无复核单",
                                    "还没有抽样复核单，请先运行【数据校验 check】。")
            return
        latest = max(htmls, key=lambda p: p.stat().st_mtime)
        try:
            webbrowser.open(latest.resolve().as_uri())
            self.statusBar().showMessage(f"已打开最新复核单：{latest.name}", 12000)
        except Exception as e:
            QMessageBox.warning(self, "打开失败", str(e))


# ---------------------------------------------------------------------------
# 深色现代主题 QSS
# ---------------------------------------------------------------------------
QSS = """
* { font-family: "Microsoft YaHei UI", "Segoe UI"; }
QWidget { background: #1b1d23; color: #e8eaf0; font-size: 14pt; }

QMainWindow, QDialog { background: #1b1d23; }
QLabel { background: transparent; }

/* 大按钮 */
QPushButton { font-size: 16pt; font-weight: 700; border: none;
              border-radius: 12px; padding: 6px 10px; color: #ffffff; }
QPushButton#btnBlue   { background: #3b82f6; }
QPushButton#btnBlue:hover   { background: #5b97ff; }
QPushButton#btnBlue:pressed { background: #2f6fd6; }
QPushButton#btnTeal   { background: #14b8a6; }
QPushButton#btnTeal:hover   { background: #2dd1bf; }
QPushButton#btnTeal:pressed { background: #0f9486; }
QPushButton#btnViolet { background: #8b5cf6; }
QPushButton#btnViolet:hover   { background: #a27fff; }
QPushButton#btnViolet:pressed { background: #7448dd; }
QPushButton#btnAmber  { background: #f59e0b; }
QPushButton#btnAmber:hover   { background: #ffb633; }
QPushButton#btnAmber:pressed { background: #d68800; }

QPushButton:disabled { background: #3a3f4b; color: #8a909e; }

QPushButton#resultButton { background: #2c313d; border: 1px solid #3d4452;
                           font-size: 15pt; font-weight: 600; border-radius: 10px; }
QPushButton#resultButton:hover { background: #383e4c; border-color: #4f8bf5; }
QPushButton#resultButton:disabled { background: #2c313d; color: #777e8c; }

QPushButton#ghostButton { background: transparent; color: #9aa1af;
                          font-size: 12pt; font-weight: 500;
                          border: 1px solid #3a404d; border-radius: 8px;
                          padding: 4px 14px; }
QPushButton#ghostButton:hover { color: #e8eaf0; border-color: #5a6272; }

/* 选项卡片 */
QFrame#optionCard { background: #24272f; border: 1px solid #343945;
                    border-radius: 10px; }
QLabel#smallLabel { color: #9aa1af; font-size: 13pt; }
QLabel#sectionTitle { font-size: 15pt; font-weight: 700; color: #eef1f7; }

/* 拖放区 */
QFrame#dropZone { background: #20242d; border: 2px dashed #4c8bf5;
                  border-radius: 12px; }
QFrame#dropZone[active="1"] { background: #262d3d;
                              border: 2px dashed #7db2ff; }
QLabel#dropTitle { font-size: 17pt; font-weight: 700; color: #eef1f7; }
QLabel#dropSub { font-size: 13pt; color: #9aa1af; }

/* 终端 */
QPlainTextEdit { background: #14161b; color: #dfe3ea;
                 border: 1px solid #2e343f; border-radius: 10px;
                 selection-background-color: #3b6fc4; }

/* 输入控件 */
QLineEdit { background: #14161b; border: 1px solid #3a404d;
            border-radius: 8px; padding: 6px 10px; font-size: 14pt;
            selection-background-color: #3b6fc4; }
QLineEdit:focus { border: 1px solid #4c8bf5; }
QLineEdit:disabled { background: #1e2128; color: #777e8c; }

QCheckBox { spacing: 8px; font-size: 14pt; color: #dfe3ea; }
QCheckBox::indicator { width: 22px; height: 22px;
                       border: 1px solid #4a5160; border-radius: 6px;
                       background: #14161b; }
QCheckBox::indicator:hover { border-color: #4c8bf5; }
QCheckBox::indicator:checked { background: #4c8bf5; border-color: #4c8bf5; }
QCheckBox:disabled { color: #777e8c; }

/* 列表 */
QListWidget { background: #14161b; border: 1px solid #2e343f;
              border-radius: 10px; padding: 6px; outline: none;
              font-size: 14pt; }
QListWidget::item { padding: 8px 10px; border-radius: 6px; color: #d5d9e2; }
QListWidget::item:hover { background: #232833; }
QListWidget::item:selected { background: #2f4a7c; color: #ffffff; }

/* 标签页 */
QTabWidget::pane { border: none; }
QTabBar { qproperty-drawBase: 0; }
QTabBar::tab { background: #24272f; color: #9aa1af;
               padding: 10px 26px; margin-right: 4px;
               border-top-left-radius: 10px; border-top-right-radius: 10px;
               font-size: 14pt; font-weight: 600; }
QTabBar::tab:selected { background: #3b82f6; color: #ffffff; }
QTabBar::tab:hover:!selected { background: #30353f; color: #e8eaf0; }

/* 状态栏 */
QStatusBar { background: #191b21; color: #9aa1af; font-size: 13pt; }
QStatusBar QLabel { padding: 2px 10px; background: transparent; font-size: 13pt; }
QLabel#exitOk { color: #6ce38a; font-weight: 700; font-size: 13pt; }
QLabel#exitFail { color: #ff7b7b; font-weight: 700; font-size: 13pt; }
QLabel#wpsOk { color: #6ce38a; font-weight: 600; }
QLabel#wpsFail { color: #ff7b7b; font-weight: 600; }
QLabel#wpsChecking { color: #ffd166; font-weight: 600; }

/* 进度条 */
QProgressBar { background: #2a2f3a; border: none; border-radius: 6px; }
QProgressBar::chunk { background: #4c8bf5; border-radius: 6px; }

/* 滚动条 */
QScrollBar:vertical { background: transparent; width: 12px; margin: 4px; }
QScrollBar::handle:vertical { background: #3a404d; border-radius: 6px;
                              min-height: 40px; }
QScrollBar::handle:vertical:hover { background: #4c8bf5; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar:horizontal { background: transparent; height: 12px; margin: 4px; }
QScrollBar::handle:horizontal { background: #3a404d; border-radius: 6px;
                                min-width: 40px; }
QScrollBar::handle:horizontal:hover { background: #4c8bf5; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; }

QToolTip { background: #24272f; color: #dfe3ea;
           border: 1px solid #3d4452; padding: 6px; }
"""


def main():
    ensure_dirs()
    ensure_runtime_scripts()

    app = QApplication(_sys.argv)
    app.setStyleSheet(QSS)
    app_font = QFont("Microsoft YaHei UI", 11)
    app.setFont(app_font)

    win = MainWindow()
    win.show()
    _sys.exit(app.exec())


if __name__ == "__main__":
    main()
