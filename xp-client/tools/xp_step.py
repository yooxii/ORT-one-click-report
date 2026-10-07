# -*- coding: utf-8 -*-
"""供 ``tools/verify_alternating.ps1`` 调用的「XP 客户端单步」工具。

每一步都走 XP 客户端自己的数据层与编辑服务（与界面同一条代码路径），
但输出一律 **ASCII**：本机控制台是 cp950，直接打印中文会抛 ``UnicodeEncodeError``，
而 PowerShell 侧要按输出做断言。中文内容不参与断言，需要比对的值由 .NET 侧写成 ASCII。

用法::

    python tools/xp_step.py --data-folder <目录> [--timeout 秒] [--retries 次] <命令> [参数]

命令：

- ``plan-exists <JobNo>``            有该计划 → ``yes``
- ``plan-remark <JobNo>``            打印计划的 Remark
- ``plan-status-is-close <JobNo>``   完成状况是 Close → ``yes``
- ``plan-count``                     计划行数
- ``requisition-exists <No>``        有该领退单 → ``yes``
- ``new-requisition <No>``           通过领退编辑服务新增（报废去向，无需回线RT）
- ``update-requisition <No> <备注>``  改备注并写变更日志（走仓储 update）
- ``changelog-count <片段>``         变更日志摘要含该片段的行数
- ``integrity``                      ``PRAGMA integrity_check`` 的结果

``new-requisition`` 失败时打印 ``locked``（数据库忙）或 ``fail`` + 具体错误，退出码仍为 0，
由调用方按输出断言。
"""

import argparse
import datetime
import os
import sys

_HERE = os.path.dirname(os.path.abspath(__file__))
_PARENT = os.path.dirname(_HERE)
if _PARENT not in sys.path:
    sys.path.insert(0, _PARENT)

from ort_xp import compat, context as context_module, logging_setup  # noqa: E402
from ort_xp.services import plan_rules  # noqa: E402

OPERATOR = "alt-verify"


def say(text):
    """只输出 ASCII（中文转义），确保任何代码页下都不会崩在打印上。"""
    line = "" if text is None else str(text)
    sys.stdout.write(line.encode("ascii", "backslashreplace").decode("ascii") + "\n")
    sys.stdout.flush()


def build_context(options):
    # 先把控制台日志关掉：AppContext 建日志时看到已有处理器就直接复用，
    # 否则 INFO 行会和命令输出混在一起，PowerShell 侧没法断言。
    logging_setup.setup_logging(compat.app_dir(), console=False)
    context = context_module.AppContext(data_folder=options.data_folder)
    if options.timeout is not None:
        context.database.timeout_seconds = int(options.timeout)
    if options.retries is not None:
        context.database.retries = int(options.retries)
    context.open()
    return context


# --------------------------------------------------------------------------- 各命令


def cmd_plan_exists(context, args):
    row = context.database.query_one('SELECT "Id" FROM plans WHERE "JobNo" = ?', (args[0],))
    say("yes" if row is not None else "no")


def cmd_plan_remark(context, args):
    row = context.database.query_one('SELECT "Remark" FROM plans WHERE "JobNo" = ?', (args[0],))
    say("(missing)" if row is None else (row["Remark"] or ""))


def cmd_plan_status_is_close(context, args):
    row = context.database.query_one('SELECT "Status" FROM plans WHERE "JobNo" = ?', (args[0],))
    if row is None:
        say("missing")
        return
    say("yes" if (row["Status"] or "").strip().lower() == "close" else "no")


def cmd_plan_count(context, args):
    say(context.database.scalar('SELECT COUNT(*) FROM "plans"', default=0))


def cmd_requisition_exists(context, args):
    row = context.database.query_one('SELECT "Id" FROM requisitions WHERE "RequisitionNo" = ?', (args[0],))
    say("yes" if row is not None else "no")


def cmd_new_requisition(context, args):
    service = context.requisition_service.with_operator(OPERATOR)
    result = service.save(
        {
            "RequisitionDate": datetime.datetime(2026, 9, 1),
            "RequisitionNo": args[0],
            "ModelName": "ALTMODEL",
            "OutQty": "2",
            "Rev": "A",
            "WorkOrder": "ALT-WO",
            "Disposition": plan_rules.DISPOSITION_SCRAP,
            "ReturnRtOrder": None,
            "Remark": "ALT-VERIFY",
        }
    )
    if result.ok:
        say("ok %s" % result.record_id)
        return
    text = " | ".join(result.errors)
    lowered = text.lower()
    say("locked" if ("lock" in lowered or "busy" in lowered) else "fail")
    say(text)


def cmd_update_requisition(context, args):
    row = context.database.query_one('SELECT "Id" FROM requisitions WHERE "RequisitionNo" = ?', (args[0],))
    if row is None:
        say("missing")
        return
    changed = context.repositories.requisitions.update(row["Id"], {"Remark": args[1]}, OPERATOR)
    say("updated" if changed else "unchanged")


def cmd_changelog_count(context, args):
    if not context.database.table_exists("plan_change_logs"):
        say(0)
        return
    say(
        context.database.scalar(
            'SELECT COUNT(*) FROM "plan_change_logs" WHERE "Summary" LIKE ?',
            ("%" + args[0] + "%",),
            default=0,
        )
    )


def cmd_integrity(context, args):
    row = context.database.query_one("PRAGMA integrity_check")
    say("ok" if row is not None and str(row[0]).lower() == "ok" else (row[0] if row is not None else "unknown"))


COMMANDS = {
    "plan-exists": cmd_plan_exists,
    "plan-remark": cmd_plan_remark,
    "plan-status-is-close": cmd_plan_status_is_close,
    "plan-count": cmd_plan_count,
    "requisition-exists": cmd_requisition_exists,
    "new-requisition": cmd_new_requisition,
    "update-requisition": cmd_update_requisition,
    "changelog-count": cmd_changelog_count,
    "integrity": cmd_integrity,
}


def main(argv=None):
    parser = argparse.ArgumentParser(prog="xp_step", description="XP 客户端单步（供交替读写验证脚本调用）")
    parser.add_argument("--data-folder", required=True, help="数据文件夹（验证脚本会用临时副本）")
    parser.add_argument("--timeout", help="覆盖数据层的 busy_timeout（秒），验证锁行为用")
    parser.add_argument("--retries", help="覆盖数据层的重试次数")
    parser.add_argument("command", help="命令名：" + ", ".join(sorted(COMMANDS)))
    parser.add_argument("args", nargs="*", help="命令参数")
    options = parser.parse_args(argv)

    if options.command not in COMMANDS:
        parser.error("unknown command: %s" % options.command)

    context = build_context(options)
    try:
        COMMANDS[options.command](context, options.args)
    finally:
        context.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
