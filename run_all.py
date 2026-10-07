# -*- coding: utf-8 -*-
"""校验一键运行：check（异构校验）→ verify（渲染闸门）。

main.py 已单独跑完、成品已生成后，用本脚本一次性完成两层校验，
省去手动依次运行 check.py 和 verify.py 的麻烦。

设计原则：
  - check 和 verify 仍可独立运行（调试时很重要），本脚本只做编排。
  - 用 subprocess 隔离调用，一个环节崩溃不影响其他环节的日志与退出码。
  - check 失败仍继续 verify，让用户一次性看到数据校验与渲染校验的完整结果。

用法：
    python run_all.py                  # check → verify（附页模式）
    python run_all.py --full-text      # check → verify（全量文本层模式）
    python run_all.py --skip-verify    # 只跑 check
    python run_all.py --skip-check     # 只跑 verify
    python run_all.py 6-8月份结算      # 批次名透传给 verify（check 不支持批次过滤）

退出码：0 = 全部通过；非 0 = 任一环节失败。
"""

import subprocess
import sys
import os

# 本脚本的开关参数（不透传给子脚本）
_LOCAL_FLAGS = {"--skip-check", "--skip-verify", "--full-text"}


def _run_step(name: str, cmd: list, extra_env: dict = None) -> int:
    """运行一个子步骤，实时透传输出，返回退出码。"""
    print("\n" + "=" * 60)
    print(f"  开始执行: {name}")
    print(f"  命令: {' '.join(cmd)}")
    print("=" * 60 + "\n")
    # 显式传递标准句柄（不能依赖默认继承）：
    # Windows 上当父进程是 pythonw.exe（无控制台，如 GUI 经 QProcess 启动本脚本）时，
    # subprocess.run 默认 close_fds=True，若此处不显式传 stdio，孙进程 check.py/verify.py
    # 的标准句柄会是 None，导致其第一次 print 即 AttributeError、静默退出（退出码 1，
    # 且留下 0 字节 _pending 日志）。显式传入父进程 stdio 后，子进程经 DuplicateHandle
    # 拿到有效管道句柄，输出可实时透传到 GUI；父进程 stdio 为 None 时退化为 DEVNULL，
    # 业务输出仍完整写入各自的日志文件，不会崩溃。
    cwd = os.path.dirname(os.path.abspath(__file__))
    env = os.environ.copy()
    if extra_env:
        env.update(extra_env)
    result = subprocess.run(
        cmd,
        cwd=cwd,
        env=env,
        stdin=sys.stdin if sys.stdin is not None else subprocess.DEVNULL,
        stdout=sys.stdout if sys.stdout is not None else subprocess.DEVNULL,
        stderr=sys.stderr if sys.stderr is not None else subprocess.DEVNULL,
    )
    code = result.returncode
    status = "通过" if code == 0 else f"失败（退出码 {code}）"
    print(f"\n  {name} 结束: {status}")
    return code


def main() -> int:
    args = sys.argv[1:]

    skip_check = "--skip-check" in args
    skip_verify = "--skip-verify" in args
    full_text = "--full-text" in args

    # 其余参数（如批次名）透传给 verify；check 不支持命令行参数
    passthrough = [a for a in args if a not in _LOCAL_FLAGS]

    results = {}
    py = sys.executable

    # 两个环节都执行时，让 check 解压一次、verify 直接复用，避免同一批 zip 解压两遍
    # （约省一半磁盘 IO 与一个 140MB 临时目录）。单独执行任一环节则不启用，各自独立解压清理。
    share_extract = (not skip_check) and (not skip_verify)
    shared_tag = "_已解压_runall"
    producer_env = {"RAWREC_EXTRACT_TAG": shared_tag, "RAWREC_EXTRACT_ROLE": "producer"}
    reuser_env = {"RAWREC_EXTRACT_TAG": shared_tag, "RAWREC_EXTRACT_ROLE": "reuser"}

    # ---- 1. check.py 异构校验 ----
    if skip_check:
        print("[跳过] check.py（--skip-check）")
    else:
        code = _run_step("check.py 异构校验", [py, "check.py"],
                         extra_env=producer_env if share_extract else None)
        results["check"] = code
        # check 失败不停止，继续 verify（渲染问题与数据问题可能独立存在）

    # ---- 2. verify.py 渲染闸门 ----
    if skip_verify:
        print("[跳过] verify.py（--skip-verify）")
    else:
        verify_cmd = [py, "verify.py"]
        if full_text:
            verify_cmd.append("--full-text")
        verify_cmd.extend(passthrough)
        code = _run_step("verify.py 渲染闸门", verify_cmd,
                         extra_env=reuser_env if share_extract else None)
        results["verify"] = code

    # ---- 汇总 ----
    print("\n" + "=" * 60)
    print("  校验汇总")
    print("=" * 60)
    for name, code in results.items():
        status = "通过" if code == 0 else f"失败（退出码 {code}）"
        print(f"  {name}: {status}")
    print("=" * 60)

    has_error = any(c != 0 for c in results.values())
    if has_error:
        print("\n存在未通过的环节，请查看各环节日志。")
        return 1
    print("\n全部通过。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
