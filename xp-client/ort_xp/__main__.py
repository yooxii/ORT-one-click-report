# -*- coding: utf-8 -*-
"""命令行入口：``python -m ort_xp``。

用法：
    python -m ort_xp                     # 启动界面
    python -m ort_xp --selftest          # 无界面自检（部署后先跑这个）
    python -m ort_xp --ui-smoke          # 界面装配自检（造窗口，不进入界面）
    python -m ort_xp --check-plans-email # 计划到期提醒演练（不发信）
    python -m ort_xp --check-plans-email --send   # 真的发信（需谨慎）
    python -m ort_xp --data-folder "Z:\\ORTData"  # 指定数据文件夹（跳过首次选择）
    python -m ort_xp --no-dialog         # 出错只写日志、不弹窗（脚本/运维用）

打包后的 exe 是**窗口子系统**：没有控制台，``sys.stderr`` 是 ``None``，往 stderr 写东西
等于什么都没写。所以这里所有启动期失败都走 :func:`ort_xp.fatal.report`（写日志 + 弹窗），
找不到数据库时还会让用户直接选数据文件夹（见 ``ui/first_run.py``）。

无界面命令（``--selftest`` / ``--ui-smoke`` / ``--check-plans-email``）的输出除了打到
stdout，也**同时写进 ``Logs\\ort_xp.log``** —— 双击启动时没有控制台，日志才是能看到的现场。
"""

import argparse
import os
import sys

from . import compat, context as context_module, fatal, logging_setup
from .version import APP_NAME, VERSION, BUILD_STAGE

EXIT_OK = 0
EXIT_PROBLEM = 1
EXIT_FATAL = 2


def build_parser():
    parser = argparse.ArgumentParser(prog="ort_xp", description="%s · XP 精简客户端" % APP_NAME)
    parser.add_argument("--data-folder", help="数据文件夹（默认按 local_settings.json / 环境变量解析）")
    parser.add_argument("--app-dir", help="程序目录（默认自动识别）")
    parser.add_argument("--selftest", action="store_true", help="无界面自检并退出")
    parser.add_argument("--ui-smoke", action="store_true", help="界面装配自检（造窗口后退出）")
    parser.add_argument("--check-plans-email", action="store_true", help="执行计划到期提醒检查")
    parser.add_argument("--send", action="store_true", help="配合 --check-plans-email 时真正发送邮件")
    parser.add_argument("--no-dialog", action="store_true", help="出错只写日志，不弹提示框（脚本/运维用）")
    parser.add_argument("--version", action="store_true", help="打印版本并退出")
    return parser


# --------------------------------------------------------------------------- 启动期提示文本


def data_folder_advice(context):
    """「数据库在哪」的四种解决办法（日志与弹窗共用）。"""
    return "\n".join(
        [
            "请让数据文件夹里有一个主程序生成的 ort_plans.db：",
            "  1) 把 ort_plans.db 复制到 %s" % os.path.join(context.app_directory, "Data"),
            "  2) 在 %s 里写 {\"DataFolder\": \"含库的目录\"}" % context.local_settings_path,
            "  3) 带参数启动：ORT-XP.exe --data-folder \"含库的目录\"",
            "  4) 重新启动，在弹出的对话框里点「是」直接选择数据文件夹",
        ]
    )


def missing_database_text(context):
    return "\n".join(
        [
            "没有找到数据库文件，程序无法启动。",
            "",
            "  数据文件夹：%s（来源：%s）" % (context.data_dir, context.data_source),
            "  数据库文件：%s（不存在）" % context.db_path,
            "",
            data_folder_advice(context),
        ]
    )


def open_failure_text(context, exc):
    return "\n".join(
        [
            "数据库打不开，程序无法启动。",
            "",
            "  原因：%s" % exc,
            "  数据文件夹：%s（来源：%s）" % (context.data_dir, context.data_source),
            "  数据库文件：%s" % context.db_path,
            "",
            data_folder_advice(context),
        ]
    )


def choose_data_folder(context, args, message):
    """数据库缺失时尝试让用户选一个数据文件夹。

    返回：目录字符串（选到了）／``""``（用户看过说明后放弃）／``None``（没能弹窗，
    调用方改用 :func:`fatal.report` 走日志 + 原生消息框）。
    """
    if args.data_folder:
        # 明确用 --data-folder 指定的路径就按它报错，不再弹选择框（脚本/运维场景）
        return None
    if not fatal.dialogs_enabled():
        return None
    try:
        from .ui import first_run
    except ImportError:
        return None
    return first_run.prompt_data_folder(context, message)


def main(argv=None):
    args = build_parser().parse_args(argv)

    headless = args.selftest or args.check_plans_email or args.ui_smoke
    if args.no_dialog or headless:
        # 无界面模式下弹窗只会把脚本卡住；现场照样进日志
        fatal.set_dialogs_enabled(False)

    if args.version:
        compat.say("%s v%s（%s）" % (APP_NAME, VERSION, BUILD_STAGE))
        return EXIT_OK

    # 日志尽早建立：窗口子系统没有控制台，日志是唯一的现场
    logger = logging_setup.setup_logging(args.app_dir or compat.app_dir())

    ok, message = compat.check_interpreter()
    compat.say(message)
    if not ok:
        return fatal.report("解释器版本不支持", message, args.app_dir, logger)

    context = context_module.AppContext(
        app_directory=args.app_dir, data_folder=args.data_folder, logger=logger
    )

    # 路径检查：Windows 上模块加载器会用 ANSI 代码页编码路径，中文（尤其是本机 cp950 编不了的
    # 简体字）会让 tkinter/PyInstaller 加载失败，而且报错很难懂。这里提前给一句人话提示。
    if not compat.can_encode_path(context.app_directory):
        message = (
            "程序所在的目录含有当前系统编码（%s）无法表示的字符：\n%s\n\n"
            "请把整个程序目录放到纯英文/数字路径下再运行，例如 C:\\ORT-XP。"
            % (compat.filesystem_encoding(), context.app_directory)
        )
        compat.say_err(message)
        if headless:
            logger.error(message)
            return EXIT_FATAL
        return fatal.report("程序路径不支持", message, context.app_directory, logger)

    logger.info("启动：%s v%s（%s）" % (APP_NAME, VERSION, BUILD_STAGE))
    logger.info("程序目录：%s；数据文件夹：%s（来源：%s）" % (context.app_directory, context.data_dir, context.data_source))

    if args.selftest:
        try:
            context.open()
        except Exception as exc:
            logger.error("自检无法连接数据库：%s" % exc)
            compat.say("自检无法连接数据库：%s" % exc)
        text = context.selftest_text()
        compat.say(text)
        # 窗口子系统下 stdout 可能是 None（双击启动时），结果必须同时留在日志里
        logger.info("自检结果：\n%s" % text)
        return EXIT_OK if context.selftest_ok() else EXIT_PROBLEM

    # 数据库文件必须先存在：sqlite 会「顺手」建一个空库，那样只会以一堆看不懂的错误收场
    if not context.database.exists():
        message = missing_database_text(context)
        choice = choose_data_folder(context, args, message)
        if choice is None:
            return fatal.report("没有找到数据库", message, context.app_directory, logger)
        if not choice:
            # 用户已经看过对话框里的说明并选择放弃，只记日志，不再弹第二个框
            logger.error(message)
            return EXIT_FATAL
        context.save_data_folder(choice)
        logger.info("已把数据文件夹写入本机设置：%s（%s）" % (choice, context.local_settings_path))

    try:
        context.open()
    except Exception as exc:
        message = open_failure_text(context, exc)
        return fatal.report("数据库打不开", message, context.app_directory, logger)

    if args.ui_smoke:
        try:
            from .ui import app as ui_app
        except ImportError as exc:
            message = "界面依赖不可用（tkinter）：%s" % exc
            compat.say_err(message)
            logger.error(message)
            context.close()
            return EXIT_FATAL
        passed, text = ui_app.ui_smoke(context)
        compat.say(text)
        logger.info("界面装配自检：%s\n%s" % ("通过" if passed else "失败", text))
        context.close()
        return EXIT_OK if passed else EXIT_PROBLEM

    if args.check_plans_email:
        text = context_module.run_reminder(context, dry_run=not args.send)
        compat.say(text)
        logger.info("计划到期提醒：\n%s" % text)
        context.close()
        return EXIT_OK

    try:
        from .ui import app as ui_app
    except ImportError as exc:
        message = "界面依赖不可用（tkinter）：%s" % exc
        compat.say_err(message)
        context.close()
        return fatal.report("缺少界面依赖", message, context.app_directory, logger)

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
