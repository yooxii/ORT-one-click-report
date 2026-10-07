# -*- coding: utf-8 -*-
"""命令行入口：``python -m ort_xp``。

用法：
    python -m ort_xp                     # 启动界面
    python -m ort_xp --selftest          # 无界面自检（部署后先跑这个）
    python -m ort_xp --check-plans-email # 计划到期提醒演练（不发信）
    python -m ort_xp --check-plans-email --send   # 真的发信（需谨慎）
"""

import argparse
import sys

from . import compat, context as context_module, logging_setup
from .version import APP_NAME, VERSION, BUILD_STAGE

EXIT_OK = 0
EXIT_PROBLEM = 1
EXIT_FATAL = 2


def build_parser():
    parser = argparse.ArgumentParser(prog="ort_xp", description="%s · XP 精简客户端" % APP_NAME)
    parser.add_argument("--data-folder", help="数据文件夹（默认按 local_settings.json / 环境变量解析）")
    parser.add_argument("--app-dir", help="程序目录（默认自动识别）")
    parser.add_argument("--selftest", action="store_true", help="无界面自检并退出")
    parser.add_argument("--check-plans-email", action="store_true", help="执行计划到期提醒检查")
    parser.add_argument("--send", action="store_true", help="配合 --check-plans-email 时真正发送邮件")
    parser.add_argument("--version", action="store_true", help="打印版本并退出")
    return parser


def main(argv=None):
    args = build_parser().parse_args(argv)

    if args.version:
        compat.say("%s v%s（%s）" % (APP_NAME, VERSION, BUILD_STAGE))
        return EXIT_OK

    ok, message = compat.check_interpreter()
    compat.say(message)
    if not ok:
        return EXIT_FATAL

    context = context_module.AppContext(app_directory=args.app_dir, data_folder=args.data_folder)
    logging_setup.get_logger().info("启动：%s v%s（%s）" % (APP_NAME, VERSION, BUILD_STAGE))

    if args.selftest:
        try:
            context.open()
        except Exception as exc:
            compat.say("自检无法连接数据库：%s" % exc)
        compat.say(context.selftest_text())
        return EXIT_OK if context.selftest_ok() else EXIT_PROBLEM

    try:
        context.open()
    except Exception as exc:
        compat.say_err("无法打开数据库：%s" % exc)
        return EXIT_FATAL

    if args.check_plans_email:
        compat.say(context_module.run_reminder(context, dry_run=not args.send))
        context.close()
        return EXIT_OK

    try:
        from .ui import app as ui_app
    except ImportError as exc:
        compat.say_err("界面依赖不可用（tkinter）：%s" % exc)
        context.close()
        return EXIT_FATAL

    try:
        return ui_app.run(context)
    finally:
        context.close()


if __name__ == "__main__":
    sys.exit(main())
