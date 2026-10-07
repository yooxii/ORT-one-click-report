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

    # 路径检查：Windows 上模块加载器会用 ANSI 代码页编码路径，中文（尤其是本机 cp950 编不了的
    # 简体字）会让 tkinter/PyInstaller 加载失败，而且报错很难懂。这里提前给一句人话提示。
    if not compat.can_encode_path(context.app_directory):
        message = (
            "程序所在的目录含有当前系统编码（%s）无法表示的字符：\n%s\n\n"
            "请把整个程序目录放到纯英文/数字路径下再运行，例如 C:\\ORT-XP。"
            % (compat.filesystem_encoding(), context.app_directory)
        )
        compat.say_err(message)
        if args.selftest or args.check_plans_email:
            return EXIT_FATAL
        try:
            from . import fatal

            fatal.show_message("程序路径不支持", message)
        except Exception:
            pass
        return EXIT_FATAL

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
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except Exception:
        from . import fatal

        sys.exit(fatal.handle(*sys.exc_info()))
