# -*- coding: utf-8 -*-
"""环境自检：解释器、标准库依赖、语法下限、数据目录与数据库、DPAPI、邮件配置。

用法::

    python tools/check_env.py
    python tools/check_env.py --data-folder "D:\\source\\repos\\ORT一键报告\\bin\\Debug\\Data"

退出码：0 = 全部通过；1 = 有失败项；2 = 连应用上下文都建不起来。
"""

import argparse
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
_PARENT = os.path.dirname(_HERE)
if _PARENT not in sys.path:
    sys.path.insert(0, _PARENT)

from ort_xp import compat, context as context_module  # noqa: E402
from ort_xp.version import APP_NAME, VERSION, BUILD_STAGE, TARGET_OS, TARGET_PYTHON  # noqa: E402
import compat_check  # noqa: E402


def main(argv=None):
    parser = argparse.ArgumentParser(description="%s · 环境自检" % APP_NAME)
    parser.add_argument("--data-folder", help="数据文件夹（默认自动解析）")
    parser.add_argument("--app-dir", help="程序目录（默认自动识别）")
    parser.add_argument("--skip-source-check", action="store_true", help="跳过 Python 3.4 语法下限检查")
    args = parser.parse_args(argv)

    compat.say("=" * 68)
    compat.say("%s v%s" % (APP_NAME, VERSION))
    compat.say("构建阶段：%s" % BUILD_STAGE)
    compat.say("目标机：%s + Python %s" % (TARGET_OS, TARGET_PYTHON))
    compat.say("=" * 68)

    failures = 0

    compat.say("")
    compat.say("[1] Python 3.4 语法下限检查")
    if args.skip_source_check:
        compat.say("    已跳过")
    else:
        findings = compat_check.check_path(os.path.join(_PARENT, "ort_xp"))
        for filename, lineno, message in findings[:40]:
            compat.say("    %s:%d: %s" % (os.path.relpath(filename, _PARENT), lineno, message))
        compat.say("    问题数：%d" % len(findings))
        if findings:
            failures += 1

    compat.say("")
    compat.say("[2] 运行环境")
    try:
        context = context_module.AppContext(app_directory=args.app_dir, data_folder=args.data_folder)
    except Exception as exc:
        compat.say("    无法建立应用上下文：%s" % exc)
        return 2

    try:
        context.open()
        opened = True
    except Exception as exc:
        compat.say("    数据库未打开：%s" % exc)
        opened = False

    compat.say("    解释器：Python %s（%s 位）" % (compat.python_version_text(), compat.python_bits()))
    compat.say("    操作系统：%s" % compat.os_description())
    compat.say("    数据目录：%s（来源：%s）" % (context.data_dir, context.data_source))
    compat.say("    数据库文件：%s" % context.db_path)

    compat.say("")
    compat.say("[3] 应用自检")
    checks = context.selftest()
    for name, ok, detail in checks:
        compat.say("    [%s] %s：%s" % ("通过" if ok else "失败", name, detail))
        if not ok:
            failures += 1

    compat.say("")
    compat.say("[4] 结论")
    if not opened:
        compat.say("    数据库不可用：先把数据文件夹指对（界面/--data-folder/local_settings.json）")
    if failures == 0:
        compat.say("    全部通过。开发环境可用；发布包仍需在 Python 3.4 下构建。")
    else:
        compat.say("    有 %d 项未通过，请按上面的说明处理后重跑。" % failures)
    compat.say("")

    if opened:
        context.close()
    return 0 if failures == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
