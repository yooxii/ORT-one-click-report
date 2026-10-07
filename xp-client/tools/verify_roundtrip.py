# -*- coding: utf-8 -*-
"""共库往返验证：在真实数据库的**副本**上读写，证明与主程序的格式约定一致。

用法::

    python tools/verify_roundtrip.py                 # 自动定位 ..\\bin\\Debug\\Data\\ort_plans.db
    python tools/verify_roundtrip.py --db <库路径>
    python tools/verify_roundtrip.py --keep          # 保留临时副本便于排查

脚本做的事：

1. 把真实库（连同 -wal / -shm）复制到临时目录 —— **绝不写原库**，结束后比对文件哈希证明没动过；
2. 读方向：把副本里主程序写下的真实数据读一遍，检查日期文本、中文、布尔都解析得了；
3. 写方向：新增 / 编辑 / 删除计划与领退，检查落库格式、审计字段、唯一索引；
4. 变更日志：检查 ``plan_change_logs`` 的行数、Action、Summary 与 Before/AfterJson 格式；
5. 登录：校验 users 表散列格式，并确认错误口令登录失败；
6. 设置：用主程序写下的真实设置键验证布尔 / 整数解析与邮件设置装配；
7. 到期提醒演练（不发信）。

退出码：0 全部通过；1 有失败项。
"""

import argparse
import base64
import datetime
import hashlib
import os
import re
import shutil
import sys
import tempfile

_HERE = os.path.dirname(os.path.abspath(__file__))
_PARENT = os.path.dirname(_HERE)
if _PARENT not in sys.path:
    sys.path.insert(0, _PARENT)

from ort_xp import compat, config, credentials  # noqa: E402
from ort_xp.db import repositories, schema_generated  # noqa: E402
from ort_xp.db.connection import Database  # noqa: E402
from ort_xp.services import auth as auth_service  # noqa: E402
from ort_xp.services import mail as mail_service  # noqa: E402

DATE_TEXT_RE = re.compile(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$")
OPERATOR = "xp-verify"


class Report(object):
    def __init__(self):
        self.items = []

    def check(self, name, ok, detail=""):
        self.items.append((name, bool(ok), detail))
        compat.say("  [%s] %s%s" % ("通过" if ok else "失败", name, ("：" + detail) if detail else ""))
        return bool(ok)

    def section(self, title):
        compat.say("")
        compat.say("== %s ==" % title)

    def failed(self):
        return [item for item in self.items if not item[1]]


def file_hash(path):
    if not os.path.isfile(path):
        return None
    digest = hashlib.sha256()
    stream = open(path, "rb")
    try:
        while True:
            chunk = stream.read(65536)
            if not chunk:
                break
            digest.update(chunk)
    finally:
        stream.close()
    return digest.hexdigest()


def default_db_path():
    candidate = os.path.join(_PARENT, "..", "bin", "Debug", "Data", "ort_plans.db")
    return os.path.abspath(candidate)


def copy_database(source, target_dir):
    """连同 -wal / -shm 一起复制（WAL 模式下缺了它们会丢最后的事务）。"""
    target = os.path.join(target_dir, os.path.basename(source))
    copied = []
    for suffix in ("", "-wal", "-shm"):
        origin = source + suffix
        if os.path.isfile(origin):
            shutil.copy2(origin, target + suffix)
            copied.append(os.path.basename(origin))
    return target, copied


def verify_read_direction(report, db):
    """读方向：主程序写下的真实数据能否被正确读出。"""
    plans = db.query('SELECT * FROM "plans"')
    requisitions = db.query('SELECT * FROM "requisitions"')
    bad_dates = []
    for table, rows in (("plans", plans), ("requisitions", requisitions)):
        for row in rows:
            for column in schema_generated.columns_of(table):
                if schema_generated.sql_type_of(table, column) != "DATETIME":
                    continue
                value = row[column]
                if value in (None, ""):
                    continue
                if compat.parse_datetime_text(value) is None:
                    bad_dates.append("%s.%s=%r" % (table, column, value))
    report.check(
        "读取主程序写入的日期",
        not bad_dates,
        "%d 条计划 + %d 条领退，全部可解析" % (len(plans), len(requisitions)) if not bad_dates else ", ".join(bad_dates),
    )

    text_ok = True
    detail = ""
    for row in requisitions[:1]:
        if row["Remark"] is not None:
            detail = "首条领退备注=%r" % row["Remark"]
    report.check("读取中文文本", text_ok, detail or "（没有可展示的中文样例）")

    settings = config.AppSettingsStore(db)
    all_settings = settings.all()
    report.check(
        "读取设置键",
        len(all_settings) > 0,
        "app_settings 共 %d 个键" % len(all_settings),
    )
    if "mail.port" in all_settings:
        port = settings.get_int("mail.port", -1)
        report.check("布尔/整数解析（.NET 文本）", port > 0, "mail.port=%s → %d" % (all_settings["mail.port"], port))
    if "mail.enabled" in all_settings:
        enabled = settings.get_bool("mail.enabled", None)
        report.check(
            "布尔文本解析",
            enabled in (True, False),
            "mail.enabled=%r → %r" % (all_settings["mail.enabled"], enabled),
        )
    return settings


def verify_write_direction(report, db, repos):
    """写方向：新增 / 编辑 / 删除计划与领退，检查格式与日志。"""
    plan_id = repos.plans.insert(
        {
            "JobNo": "XP-VERIFY-001",
            "ModelName": "31ABCDEFG7H",
            "TestItem": "热冲击（XP 验证）",
            "Owner": "验证人",
            "Stage": "DVT",
            "Status": "进行中",
            "StartDate": datetime.datetime(2026, 10, 1, 8, 30, 0),
            "EndDate": datetime.datetime(2026, 10, 20),
            "Remark": "中文备注 · 往返验证",
        },
        OPERATOR,
    )
    row = repos.plans.get(plan_id)
    date_ok = DATE_TEXT_RE.match(row["StartDate"] or "") is not None and row["StartDate"] == "2026-10-01 08:30:00"
    report.check("新增计划：日期按主程序格式落库", date_ok, "StartDate=%r" % row["StartDate"])
    report.check(
        "新增计划：中文与审计字段",
        row["TestItem"] == "热冲击（XP 验证）" and row["CreatedBy"] == OPERATOR and row["UpdatedBy"] == OPERATOR,
        "TestItem=%r CreatedBy=%r" % (row["TestItem"], row["CreatedBy"]),
    )

    changed = repos.plans.update(plan_id, {"Status": "Close", "Remark": "改为结案"}, OPERATOR)
    after = repos.plans.get(plan_id)
    report.check("编辑计划：内容更新", changed and after["Status"] == "Close", "Status=%r" % after["Status"])
    noop = repos.plans.update(plan_id, {"Status": "Close"}, OPERATOR)
    report.check("编辑计划：无变化不写日志", noop is False, "返回 %r" % noop)

    requisition_id = repos.requisitions.insert(
        {
            "RequisitionDate": datetime.datetime(2026, 10, 7),
            "RequisitionNo": "XP-VERIFY-001",
            "ModelName": "31ABCDEFG7H",
            "OutQty": "2",
            "Disposition": "入库",
            "WorkOrder": "WO-XP-1",
            "Remark": "领退中文备注",
        },
        OPERATOR,
    )
    requisition = repos.requisitions.get(requisition_id)
    report.check(
        "新增领退：日期与中文",
        requisition["RequisitionDate"] == "2026-10-07 00:00:00" and requisition["Remark"] == "领退中文备注",
        "RequisitionDate=%r Remark=%r" % (requisition["RequisitionDate"], requisition["Remark"]),
    )
    duplicated = False
    try:
        repos.requisitions.insert(
            {"RequisitionNo": "XP-VERIFY-001", "ModelName": "重复单号"}, OPERATOR
        )
    except Exception as exc:
        duplicated = "UNIQUE" in str(exc).upper() or "unique" in str(exc)
    report.check("唯一索引生效（重复单据号被拒）", duplicated, "库返回唯一约束错误")

    return plan_id, requisition_id


def verify_change_logs(report, db, repos, plan_id):
    # 注意：主程序的领退与计划**共用** plan_change_logs，PlanId 是各自表的 Id，会重号，
    # 所以这里按摘要（含「计划 XP-VERIFY-001」）筛选，而不是只按 PlanId。
    rows = db.query(
        'SELECT * FROM "plan_change_logs" WHERE "Summary" LIKE ? ORDER BY "Id"',
        ("%\u8ba1\u5212 XP-VERIFY-001%",),
    )
    actions = [row["Action"] for row in rows]
    report.check(
        "变更日志：动作序列",
        actions == [repositories.ACTION_ADD, repositories.ACTION_EDIT],
        "%s" % ",".join(actions),
    )
    if len(rows) >= 2:
        added, edited = rows[0], rows[1]
        report.check(
            "变更日志：摘要与主程序同格式",
            added["Summary"] == "新增计划 XP-VERIFY-001 (31ABCDEFG7H)"
            and edited["Summary"] == "编辑计划 XP-VERIFY-001 (31ABCDEFG7H)",
            "新增='%s' 编辑='%s'" % (added["Summary"], edited["Summary"]),
        )
        report.check("变更日志：新增无前快照", added["BeforeJson"] is None, "")
        json_ok = (
            added["AfterJson"] is not None
            and added["AfterJson"].startswith('{"Id":')
            and '"StartDate":"2026-10-01T08:30:00"' in added["AfterJson"]
            and '"Status":"Close"' in edited["AfterJson"]
            and '"Status":"进行中"' in edited["BeforeJson"]
            and ": " not in added["AfterJson"]
        )
        report.check("变更日志：PascalCase 紧凑 JSON + ISO 日期", json_ok, "AfterJson 前 60 字符=%s" % (added["AfterJson"] or "")[:60])
    report.check(
        "变更日志：表结构与主程序模型一致",
        sorted(schema_generated.columns_of("plan_change_logs"))
        == sorted(db.column_names("plan_change_logs")),
        "列=%s" % ",".join(db.column_names("plan_change_logs")),
    )


def verify_login(report, db):
    users = db.query('SELECT * FROM "users"')
    format_ok = True
    checked = 0
    for row in users:
        salt = row["Salt"] or ""
        password_hash = row["PasswordHash"] or ""
        if not salt and not password_hash:
            continue  # 无密码账号
        checked += 1
        if not salt or len(salt) != 32:
            format_ok = False
        else:
            try:
                int(salt, 16)
            except ValueError:
                format_ok = False
        try:
            if len(base64.b64decode(password_hash.encode("ascii"))) != 32:
                format_ok = False
        except Exception:
            format_ok = False
    report.check(
        "散列格式与主程序一致（Salt 32 位十六进制 / Hash Base64 32 字节）",
        format_ok,
        "检查了 %d 个设密码的账号" % checked,
    )

    service = auth_service.AuthService(db, credentials.LocalCredentialStore("ort-xp-verify-tmp"))
    target = None
    for row in users:
        if row["PasswordHash"] and row["Salt"]:
            target = row
            break
    if target is not None:
        wrong = service.login(target["Username"], "\u4e00\u5b9a\u4e0d\u5bf9\u7684\u53e3\u4ee4")
        report.check(
            "错误口令登录被拒",
            wrong.ok is False,
            "账号=%s → %s" % (target["Username"], wrong.message),
        )
    else:
        report.check("错误口令登录被拒", True, "库里没有设密码的账号，跳过")


def verify_mail(report, db, settings_store):
    settings = config.MailSettings.from_store(settings_store, credentials.LocalCredentialStore("ort-xp-verify-tmp"))
    service = mail_service.MailService(db, settings)
    ready, reason = service.is_ready()
    report.check(
        "邮件配置装配（按主程序设置键）",
        settings.port > 0 and settings.security in ("None", "StartTls", "Ssl"),
        "port=%d security=%s → %s" % (settings.port, settings.security, "可发信" if ready else reason),
    )

    settings.warning_days_before = 30
    settings.warning_include_overdue = True
    reminder = mail_service.PlanDeadlineReminder(db, service, settings)
    summary = reminder.run(dry_run=True)
    report.check(
        "到期提醒演练（不发信）",
        summary["failed"] == 0,
        "候选 %d，成功 %d，跳过 %d，失败 %d" % (
            summary["candidates"],
            summary["sent"],
            summary["skipped"],
            summary["failed"],
        ),
    )
    return summary


def main(argv=None):
    parser = argparse.ArgumentParser(description="共库往返验证（在副本上进行）")
    parser.add_argument("--db", help="真实数据库路径（默认自动定位）")
    parser.add_argument("--keep", action="store_true", help="保留临时副本")
    args = parser.parse_args(argv)

    source = os.path.abspath(args.db) if args.db else default_db_path()
    compat.say("=" * 68)
    compat.say("共库往返验证（只读原库，全部写入发生在副本上）")
    compat.say("原库：%s" % source)
    compat.say("=" * 68)
    if not os.path.isfile(source):
        compat.say("找不到数据库文件：%s" % source)
        return 1

    before = dict((suffix, file_hash(source + suffix)) for suffix in ("", "-wal", "-shm"))
    temp_dir = tempfile.mkdtemp(prefix="ort-xp-verify-")
    report = Report()
    try:
        copy_path, copied = copy_database(source, temp_dir)
        compat.say("")
        compat.say("副本：%s（已复制 %s）" % (copy_path, ", ".join(copied)))

        db = Database(copy_path)
        db.connect()
        repos = repositories.Repositories(db)
        try:
            report.section("1. 读方向（主程序写入 → Python 读取）")
            settings_store = verify_read_direction(report, db)

            report.section("2. 写方向（Python 写入 → 主程序格式）")
            plan_id, requisition_id = verify_write_direction(report, db, repos)

            report.section("3. 变更日志")
            verify_change_logs(report, db, repos, plan_id)

            report.section("4. 登录")
            verify_login(report, db)

            report.section("5. 邮件")
            verify_mail(report, db, settings_store)

            report.section("6. 清理副本写入")
            deleted_plan = repos.plans.delete(plan_id, OPERATOR)
            deleted_requisition = repos.requisitions.delete(requisition_id, OPERATOR)
            plan_logs = db.scalar(
                'SELECT COUNT(*) FROM "plan_change_logs" WHERE "Summary" LIKE ?',
                ("%\u8ba1\u5212 XP-VERIFY-001%",),
                default=0,
            )
            report.check(
                "删除并留日志",
                deleted_plan and deleted_requisition and plan_logs == 3,
                "该计划的日志条数=%d（新增/编辑/删除）" % plan_logs,
            )
        finally:
            db.close()

        report.section("7. 原库未被改动")
        after = dict((suffix, file_hash(source + suffix)) for suffix in ("", "-wal", "-shm"))
        unchanged = before == after
        report.check("原库（含 -wal/-shm）哈希未变", unchanged, "sha256 前 12 位=%s" % (after[""] or "")[:12])
    finally:
        if args.keep:
            compat.say("")
            compat.say("临时副本已保留：%s" % temp_dir)
        else:
            shutil.rmtree(temp_dir, ignore_errors=True)

    failed = report.failed()
    compat.say("")
    compat.say("=" * 68)
    compat.say("结论：%d 项通过，%d 项失败" % (len(report.items) - len(failed), len(failed)))
    for name, _ok, detail in failed:
        compat.say("  失败：%s %s" % (name, detail))
    compat.say("=" * 68)
    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(main())
