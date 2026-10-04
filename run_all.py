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


def _run_step(name: str, cmd: list) -> int:
    """运行一个子步骤，实时透传输出，返回退出码。"""
    print("\n" + "=" * 60)
    print(f"  开始执行: {name}")
    print(f"  命令: {' '.join(cmd)}")
    print("=" * 60 + "\n")
    # 继承父进程 stdout/stderr，实时显示输出；cwd 固定为项目根目录
    result = subprocess.run(cmd, cwd=os.path.dirname(os.path.abspath(__file__)))
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

    # ---- 1. check.py 异构校验 ----
    if skip_check:
        print("[跳过] check.py（--skip-check）")
    else:
        code = _run_step("check.py 异构校验", [py, "check.py"])
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
        code = _run_step("verify.py 渲染闸门", verify_cmd)
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
