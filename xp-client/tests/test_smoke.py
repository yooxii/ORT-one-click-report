# -*- coding: utf-8 -*-
"""标准库 unittest（XP 上不引入第三方测试框架）。

运行::

    python -m unittest discover -s tests -v
"""

import argparse
import datetime
import logging
import os
import shutil
import sqlite3
import sys
import tempfile
import unittest

_HERE = os.path.dirname(os.path.abspath(__file__))
_PARENT = os.path.dirname(_HERE)
for _path in (_PARENT, os.path.join(_PARENT, "tools")):
    if _path not in sys.path:
        sys.path.insert(0, _path)

import compat_check  # noqa: E402
from ort_xp import __main__ as main_module  # noqa: E402
from ort_xp import compat, config, credentials, dpapi, fatal, logging_setup  # noqa: E402
from ort_xp import context as context_module  # noqa: E402
from ort_xp.db import repositories  # noqa: E402
from ort_xp.db import schema_generated  # noqa: E402
from ort_xp.db.connection import Database  # noqa: E402
from ort_xp.services import auth as auth_service  # noqa: E402
from ort_xp.services import mail as mail_service  # noqa: E402
from ort_xp.services import plan_rules  # noqa: E402
from ort_xp.services import plans as plans_service  # noqa: E402
from ort_xp.ui import first_run  # noqa: E402

#: 与主程序 AuthService.HashPassword 对应的向量：Base64(SHA256(UTF8("abc123" + "p@ss")))
PASSWORD_VECTOR_SALT = "abc123"
PASSWORD_VECTOR_PASSWORD = "p@ss"
PASSWORD_VECTOR_HASH = "w8Pg4xD0k+t2XpSppuBUjy1sKt+JIMiet5Nz09HNPgE="

SCHEMA_SQL = """
CREATE TABLE users (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    Username NVARCHAR(64) NOT NULL,
    DisplayName NVARCHAR(64) NULL,
    Email NVARCHAR(128) NULL,
    PasswordHash NVARCHAR(128) NOT NULL,
    Salt NVARCHAR(64) NOT NULL,
    IsActive BOOLEAN NOT NULL,
    CreatedAt DATETIME NOT NULL
);
CREATE TABLE user_roles (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    UserId INTEGER NOT NULL,
    Role NVARCHAR(32) NOT NULL
);
CREATE TABLE app_settings (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    Key NVARCHAR(64) NOT NULL,
    Value NVARCHAR(512) NULL
);
CREATE TABLE mail_logs (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    Kind NVARCHAR(16) NOT NULL,
    RefType NVARCHAR(32) NULL,
    RefKey NVARCHAR(128) NULL,
    Recipients NVARCHAR(1024) NULL,
    Subject NVARCHAR(512) NULL,
    Success BOOLEAN NOT NULL,
    Error NVARCHAR(1024) NULL,
    CreatedAt DATETIME NOT NULL
);
CREATE TABLE plans (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    ModelName NVARCHAR(128) NULL,
    TestItem NVARCHAR(128) NULL,
    Remark NVARCHAR(512) NULL,
    JobNo NVARCHAR(64) NULL,
    Product NVARCHAR(64) NULL,
    Customer NVARCHAR(64) NULL,
    Stage NVARCHAR(32) NULL,
    SampleSize NVARCHAR(32) NULL,
    TestPeriod NVARCHAR(32) NULL,
    Owner NVARCHAR(64) NULL,
    StartDate DATETIME NULL,
    EndDate DATETIME NULL,
    Status NVARCHAR(32) NULL,
    ReportStatus NVARCHAR(16) NULL,
    UploadELab NVARCHAR(32) NULL,
    UnitReturnDate DATETIME NULL,
    CreatedBy NVARCHAR(64) NULL,
    CreatedAt DATETIME NULL,
    UpdatedBy NVARCHAR(64) NULL,
    UpdatedAt DATETIME NULL
);
CREATE TABLE requisitions (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    RequisitionDate DATETIME NULL,
    RequisitionNo NVARCHAR(64) NULL,
    ModelName NVARCHAR(128) NULL,
    OutQty NVARCHAR(32) NULL,
    Disposition NVARCHAR(16) NULL,
    SN NVARCHAR(2048) NULL,
    SnFilePath NVARCHAR(512) NULL,
    Rev NVARCHAR(32) NULL,
    WorkOrder NVARCHAR(64) NULL,
    DC NVARCHAR(32) NULL,
    LineNo NVARCHAR(32) NULL,
    ReturnRtOrder NVARCHAR(64) NULL,
    ReturnQty NVARCHAR(32) NULL,
    ReturnDate DATETIME NULL,
    StockInNo NVARCHAR(64) NULL,
    StockInQty NVARCHAR(32) NULL,
    StockInDate DATETIME NULL,
    ScrapNo NVARCHAR(64) NULL,
    ScrapQty NVARCHAR(32) NULL,
    ScrapDate DATETIME NULL,
    ScrapSnText TEXT NULL,
    ScrapSnFilePath NVARCHAR(512) NULL,
    Remark NVARCHAR(512) NULL,
    CreatedBy NVARCHAR(64) NULL,
    CreatedAt DATETIME NULL,
    UpdatedBy NVARCHAR(64) NULL,
    UpdatedAt DATETIME NULL
);
-- 下拉取值用的字典表
CREATE TABLE stages (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    Name NVARCHAR(32) NOT NULL,
    Description NVARCHAR(256) NULL
);
CREATE TABLE code_mappings (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    CodeType NVARCHAR(1) NOT NULL,
    Code NVARCHAR(8) NOT NULL,
    Name NVARCHAR(64) NOT NULL
);
CREATE TABLE products (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    Name NVARCHAR(64) NOT NULL,
    Code NVARCHAR(32) NULL,
    Remark NVARCHAR(256) NULL
);
CREATE TABLE customers (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    Name NVARCHAR(64) NOT NULL,
    Code NVARCHAR(32) NULL,
    Remark NVARCHAR(256) NULL
);
CREATE TABLE model_mappings (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    ModelName NVARCHAR(128) NOT NULL,
    Product NVARCHAR(64) NULL,
    Customer NVARCHAR(64) NULL
);
CREATE TABLE test_items_catalog (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    Name NVARCHAR(128) NOT NULL,
    Period NVARCHAR(32) NULL,
    Category NVARCHAR(64) NULL,
    Owner NVARCHAR(64) NULL,
    OwnerIds NVARCHAR(256) NULL,
    Remark NVARCHAR(256) NULL
);
-- 唯一索引与真实库一致（FreeSql 按模型上的 [Index] 生成）
CREATE UNIQUE INDEX uk_username ON users("Username");
CREATE UNIQUE INDEX uk_job_no ON plans("JobNo");
CREATE UNIQUE INDEX uk_req_requisition_no ON requisitions("RequisitionNo");
"""


class TempDatabaseTest(unittest.TestCase):
    """建一个临时库（与真实库同名同结构的最小版本）。"""

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="ort-xp-test-")
        self.db_path = os.path.join(self.temp_dir, "ort_plans.db")
        connection = sqlite3.connect(self.db_path)
        try:
            connection.executescript(SCHEMA_SQL)
            connection.commit()
        finally:
            connection.close()
        self.db = Database(self.db_path)
        self.db.connect()

    def tearDown(self):
        self.db.close()
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def add_user(self, username, password, display_name="测试用户", email="user@example.com", active=True):
        salt = auth_service.new_salt()
        password_hash = auth_service.hash_password(salt, password) if password is not None else ""
        if password is None:
            salt = ""
        self.db.execute(
            "INSERT INTO users (Username, DisplayName, Email, PasswordHash, Salt, IsActive, CreatedAt) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            (username, display_name, email, password_hash, salt, 1 if active else 0, "2026-01-01 09:00:00"),
        )
        return self.db.scalar("SELECT Id FROM users WHERE Username = ?", (username,))


class CompatTests(unittest.TestCase):
    def test_interpreter_ok(self):
        ok, message = compat.check_interpreter()
        self.assertTrue(ok, message)

    def test_bool_text_parsing(self):
        self.assertTrue(compat.to_bool("True"))
        self.assertTrue(compat.to_bool("true"))
        self.assertTrue(compat.to_bool("1"))
        self.assertFalse(compat.to_bool("False"))
        self.assertFalse(compat.to_bool(""))
        self.assertTrue(compat.to_bool(None, True))
        self.assertEqual("True", compat.bool_text(True))
        self.assertEqual("False", compat.bool_text(False))

    def test_number_parsing(self):
        self.assertEqual(14, compat.to_int("14"))
        self.assertEqual(3, compat.to_int("3.9"))
        self.assertEqual(5, compat.to_int(None, 5))
        self.assertAlmostEqual(1.5, compat.to_float("1.5"))

    def test_datetime_text_two_precisions(self):
        first = compat.parse_datetime_text("2026-09-01 00:00:00")
        second = compat.parse_datetime_text("2026-09-29 11:29:17.806635")
        self.assertEqual((2026, 9, 1), (first.year, first.month, first.day))
        self.assertEqual(806635, second.microsecond)
        self.assertIsNone(compat.parse_datetime_text(""))
        self.assertIsNone(compat.parse_datetime_text("不是日期"))
        self.assertEqual("2026-09-01 00:00:00", compat.format_datetime_text(first))

    def test_dotnet_date_format(self):
        import datetime

        value = datetime.datetime(2026, 10, 7, 9, 5, 3)
        self.assertEqual("2026/10/7", compat.format_dotnet_date(value, "yyyy/M/d"))
        self.assertEqual("2026/10/07", compat.format_dotnet_date(value, "yyyy/MM/dd"))
        self.assertEqual("09:05:03", compat.format_dotnet_date(value, "HH:mm:ss"))
        self.assertEqual("26-10-07", compat.format_dotnet_date(value, "yy-MM-dd"))

    def test_path_encoding_guard(self):
        """Windows 下 ANSI 代码页编不了的路径要让程序提前报错，而不是崩在加载器里。"""
        self.assertTrue(compat.can_encode_path("C:\\ORT-XP"))
        self.assertTrue(compat.can_encode_path(""))
        self.assertTrue(compat.can_encode_path(None))
        encoding = compat.filesystem_encoding()
        if compat.is_windows() and encoding.lower() == "mbcs":
            # 任何 ANSI 代码页都表示不了 emoji
            self.assertFalse(compat.can_encode_path(u"D:\\\U0001F600"))
            probe = u"D:\\source\\ORT\u4e00\u952e\u62a5\u544a"  # 含简体「一键报告」
            try:
                probe.encode(encoding)
                expected = True
            except UnicodeEncodeError:
                expected = False
            self.assertEqual(expected, compat.can_encode_path(probe))
        else:
            # Python 3.6+ 的 Windows 文件系统编码是 UTF-8，中文路径本来就没问题
            self.assertTrue(compat.can_encode_path(u"D:\\ORT\u4e00\u952e\u62a5\u544a"))


class DatabaseTests(TempDatabaseTest):
    def test_table_and_column_detection(self):
        self.assertTrue(self.db.table_exists("users"))
        self.assertFalse(self.db.table_exists("plan_change_logs"))
        self.assertIn("Username", self.db.column_names("users"))
        self.assertTrue(self.db.has_column("users", "username"))
        self.assertFalse(self.db.has_column("users", "NotThere"))

    def test_temp_schema_matches_generated_schema(self):
        """临时库的列必须与从主程序模型生成的清单完全一致，避免测试与实际字段脱节。"""
        for table in ("plans", "requisitions", "users", "mail_logs", "app_settings"):
            expected = set(schema_generated.columns_of(table))
            actual = set(self.db.column_names(table))
            self.assertEqual(expected, actual, "%s 列不一致" % table)

    def test_transaction_commit_and_rollback(self):
        with self.db.transaction() as connection:
            connection.execute(
                "INSERT INTO requisitions (RequisitionNo, ModelName) VALUES (?, ?)", ("RT2601", "X100")
            )
        self.assertEqual(1, self.db.scalar("SELECT COUNT(*) FROM requisitions"))

        with self.assertRaises(RuntimeError):
            with self.db.transaction() as connection:
                connection.execute(
                    "INSERT INTO requisitions (RequisitionNo, ModelName) VALUES (?, ?)", ("RT2602", "X200")
                )
                raise RuntimeError("模拟异常")
        self.assertEqual(1, self.db.scalar("SELECT COUNT(*) FROM requisitions"))

    def test_health_reports_tables(self):
        health = self.db.health()
        self.assertIsNone(health["error"])
        self.assertIn("users", health["tables"])
        self.assertIsNotNone(health["journal_mode"])


class SettingsTests(TempDatabaseTest):
    def test_app_settings_roundtrip(self):
        store = config.AppSettingsStore(self.db)
        self.assertTrue(store.available())
        store.set("mail.enabled", True)
        store.set("mail.port", 587)
        store.set("mail.host", "smtp.example.com")
        self.assertEqual("True", store.get("mail.enabled"))
        self.assertTrue(store.get_bool("mail.enabled"))
        self.assertEqual(587, store.get_int("mail.port"))
        self.assertEqual("smtp.example.com", store.get_str("mail.host"))
        self.assertFalse(store.get_bool("mail.warningEnabled", False))
        store.set("mail.enabled", False)
        self.assertEqual(1, self.db.scalar("SELECT COUNT(*) FROM app_settings WHERE Key = ?", ("mail.enabled",)))
        self.assertFalse(store.get_bool("mail.enabled", True))

    def test_mail_settings_from_store(self):
        store = config.AppSettingsStore(self.db)
        store.set_many(
            {
                "mail.host": "smtp.example.com",
                "mail.port": 465,
                "mail.security": "Ssl",
                "mail.fromAddress": "ort@example.com",
                "mail.warningDaysBefore": 5,
                "mail.ccAdmin.Warning": True,
                "mail.template.Warning.subject": "{{JobNo}} 提醒",
            }
        )
        settings = config.MailSettings.from_store(store)
        self.assertEqual("smtp.example.com", settings.host)
        self.assertEqual(465, settings.port)
        self.assertEqual("Ssl", settings.security)
        self.assertEqual(5, settings.warning_days_before)
        self.assertTrue(settings.cc_admins.get("Warning"))
        self.assertEqual("{{JobNo}} 提醒", settings.template("Warning", True))


class LocalSettingsTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="ort-xp-cfg-")

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_roundtrip_and_resolution(self):
        path = config.local_settings_path(self.temp_dir)
        settings = config.LocalSettings(data_folder="Z:\\ORT数据")
        settings.save(path)
        loaded = config.LocalSettings.load(path)
        self.assertEqual("Z:\\ORT数据", loaded.data_folder)

        data_dir, source = config.resolve_data_dir(self.temp_dir, environ={})
        self.assertEqual("Z:\\ORT数据", data_dir)
        self.assertIn("本机设置", source)

        data_dir, source = config.resolve_data_dir(
            self.temp_dir, environ={config.ENV_DATA_FOLDER: "Y:\\另一个"}
        )
        self.assertEqual("Y:\\另一个", data_dir)
        self.assertIn("环境变量", source)

    def test_default_data_dir(self):
        data_dir, source = config.resolve_data_dir(self.temp_dir, environ={})
        self.assertEqual(os.path.join(self.temp_dir, "Data"), data_dir)
        self.assertIn("默认", source)

    def test_unc_detection(self):
        self.assertTrue(config.is_unc_path("\\\\server\\share\\Data"))
        self.assertFalse(config.is_unc_path("Z:\\Data"))


class AuthTests(TempDatabaseTest):
    def test_hash_matches_dotnet_vector(self):
        self.assertEqual(
            PASSWORD_VECTOR_HASH,
            auth_service.hash_password(PASSWORD_VECTOR_SALT, PASSWORD_VECTOR_PASSWORD),
        )

    def test_salt_shape(self):
        salt = auth_service.new_salt()
        self.assertEqual(32, len(salt))
        self.assertEqual(salt.lower(), salt)
        int(salt, 16)
        self.assertNotEqual(salt, auth_service.new_salt())

    def test_login_success_and_roles(self):
        user_id = self.add_user("alice", "secret")
        self.db.execute("INSERT INTO user_roles (UserId, Role) VALUES (?, ?)", (user_id, "Administrator"))
        service = auth_service.AuthService(self.db)
        result = service.login("alice", "secret")
        self.assertTrue(result.ok, result.message)
        self.assertEqual(["Administrator"], service.roles)
        self.assertTrue(service.has_role("administrator"))
        self.assertFalse(service.has_role("Technician"))
        self.assertTrue(service.is_authenticated)

    def test_login_wrong_password_and_inactive(self):
        self.add_user("bob", "secret")
        self.add_user("carol", "secret", active=False)
        service = auth_service.AuthService(self.db)
        self.assertFalse(service.login("bob", "wrong").ok)
        self.assertFalse(service.login("bob", "").ok)
        self.assertFalse(service.login("carol", "secret").ok)
        self.assertFalse(service.login("nobody", "secret").ok)

    def test_passwordless_login(self):
        self.add_user("dave", None)
        service = auth_service.AuthService(self.db)
        result = service.login("dave", "")
        self.assertTrue(result.ok)
        self.assertTrue(result.passwordless)

    def test_logout_clears_state(self):
        self.add_user("erin", "secret")
        service = auth_service.AuthService(self.db)
        service.login("erin", "secret")
        service.logout()
        self.assertFalse(service.is_authenticated)
        self.assertEqual([], service.roles)


class MailTests(TempDatabaseTest):
    def setUp(self):
        TempDatabaseTest.setUp(self)
        self.settings = config.MailSettings()
        self.settings.enabled = True
        self.settings.host = "smtp.example.com"
        self.settings.port = 587
        self.settings.security = "StartTls"
        self.settings.from_address = "ort@example.com"
        self.settings.warning_days_before = 3
        self.settings.warning_include_overdue = True
        self.service = mail_service.MailService(self.db, self.settings)

    def test_template_render(self):
        variables = {"JobNo": "RT2601", "Owner": "王工", "Missing": None, "DaysLeft": -2}
        text = mail_service.render_template("{{JobNo}}/{{Owner}}/{{Missing}}/{{DaysLeft}}", variables)
        self.assertEqual("RT2601/王工//-2", text)
        import datetime

        text = mail_service.render_template("{{EndDate|yyyy/M/d}}", {"EndDate": datetime.datetime(2026, 10, 7)})
        self.assertEqual("2026/10/7", text)

    def test_address_helpers(self):
        self.assertTrue(mail_service.is_valid_address("a@b.com"))
        self.assertFalse(mail_service.is_valid_address("a@b"))
        self.assertFalse(mail_service.is_valid_address(""))
        self.assertEqual(["a@b.com", "c@d.com"], mail_service.split_list("a@b.com; c@d.com"))
        self.assertEqual(["a@b.com", "c@d.com"], mail_service.split_list("a@b.com，c@d.com"))
        self.assertTrue(mail_service.is_completed("Close"))
        self.assertTrue(mail_service.is_completed(" close "))
        self.assertFalse(mail_service.is_completed("进行中"))
        self.assertFalse(mail_service.is_completed(None))

    def test_normalize_recipients(self):
        result = self.service.normalize_recipients(["A@B.com", "a@b.com", "bad", ""])
        self.assertEqual(["A@B.com"], result)

    def test_resolve_user_email(self):
        self.add_user("alice", "secret", display_name="爱丽丝", email="alice@example.com")
        self.assertEqual("alice@example.com", self.service.resolve_user_email("爱丽丝"))
        self.assertEqual("alice@example.com", self.service.resolve_user_email("alice"))
        self.assertIsNone(self.service.resolve_user_email("不存在的人"))

    def test_is_ready_reasons(self):
        self.settings.enabled = False
        self.assertFalse(self.service.is_ready()[0])
        self.settings.enabled = True
        self.settings.use_default_credentials = True
        ready, reason = self.service.is_ready()
        self.assertFalse(ready)
        self.assertIn("集成验证", reason)
        self.settings.use_default_credentials = False
        self.assertTrue(self.service.is_ready()[0])

    def test_send_dry_run_writes_no_log(self):
        result = self.service.send(
            mail_service.MAIL_KIND_WARNING, ["a@b.com"], "主题", "正文", ref_type="Plan", ref_key="RT2601", dry_run=True
        )
        self.assertTrue(result.success)
        self.assertEqual([], self.service.recent_logs())

    def test_deadline_collect_filters(self):
        import datetime

        today = datetime.datetime(2026, 10, 7, 8, 0, 0)
        rows = [
            ("RT2601", "X100", "王工", "2026-10-08 00:00:00", "进行中"),
            ("RT2602", "X200", "王工", "2026-10-01 00:00:00", "进行中"),
            ("RT2603", "X300", "王工", "2026-10-09 00:00:00", "Close"),
            ("RT2604", "X400", "王工", "2026-12-01 00:00:00", "进行中"),
        ]
        for job_no, model, owner, end_date, status in rows:
            self.db.execute(
                "INSERT INTO plans (JobNo, ModelName, Owner, EndDate, Status) VALUES (?, ?, ?, ?, ?)",
                (job_no, model, owner, end_date, status),
            )
        reminder = mail_service.PlanDeadlineReminder(self.db, self.service, self.settings)
        collected = reminder.collect(today)
        job_numbers = [item["job_no"] for item in collected]
        self.assertEqual(["RT2602", "RT2601"], job_numbers)
        overdue = [item for item in collected if item["job_no"] == "RT2602"][0]
        self.assertTrue(overdue["overdue"])
        self.assertEqual(-6, overdue["days_left"])

        self.settings.warning_include_overdue = False
        job_numbers = [item["job_no"] for item in reminder.collect(today)]
        self.assertEqual(["RT2601"], job_numbers)

    def test_deadline_run_dry_run(self):
        import datetime

        user_id = self.add_user("wang", "secret", display_name="王工", email="wang@example.com")
        self.db.execute("INSERT INTO user_roles (UserId, Role) VALUES (?, ?)", (user_id, "Technician"))
        self.db.execute(
            "INSERT INTO plans (JobNo, ModelName, Owner, EndDate, Status) VALUES (?, ?, ?, ?, ?)",
            ("RT2601", "X100", "王工", "2026-10-08 00:00:00", "进行中"),
        )
        reminder = mail_service.PlanDeadlineReminder(self.db, self.service, self.settings)
        summary = reminder.run(dry_run=True, today=datetime.datetime(2026, 10, 7, 8, 0, 0))
        self.assertEqual(1, summary["candidates"])
        self.assertEqual(1, summary["sent"])
        self.assertEqual(0, summary["failed"])


@unittest.skipUnless(dpapi.AVAILABLE, "DPAPI 不可用")
class CredentialTests(unittest.TestCase):
    FOLDER = "ort-xp-selftest-tmp"

    def setUp(self):
        self.store = credentials.LocalCredentialStore(self.FOLDER)
        self.store.clear_login()
        self.store.clear_mail_password()

    def tearDown(self):
        self.store.clear_login()
        self.store.clear_mail_password()
        shutil.rmtree(self.store.base_dir, ignore_errors=True)

    def test_login_credentials_roundtrip(self):
        self.assertTrue(self.store.save_login("alice", "secret", days=1))
        loaded = self.store.load_login()
        self.assertEqual(("alice", "secret"), loaded)
        self.store.clear_login()
        self.assertIsNone(self.store.load_login())

    def test_expired_credential_is_dropped(self):
        self.assertTrue(self.store.save_login("alice", "secret", days=-1))
        self.assertIsNone(self.store.load_login())

    def test_mail_password_roundtrip(self):
        self.assertTrue(self.store.save_mail_password("smtp-secret"))
        self.assertEqual("smtp-secret", self.store.load_mail_password())
        self.store.clear_mail_password()
        self.assertIsNone(self.store.load_mail_password())


class RepositoryTests(TempDatabaseTest):
    """数据层读写：行内容、日期格式、变更日志、唯一性、下拉取值。"""

    OPERATOR = "tester"

    def setUp(self):
        TempDatabaseTest.setUp(self)
        self.repos = repositories.Repositories(self.db)

    def plan_values(self, **overrides):
        values = {
            "JobNo": "RT2601-001",
            "ModelName": "31ABCDEFG7H",
            "TestItem": "热冲击",
            "Owner": "王工",
            "Stage": "DVT",
            "Status": "进行中",
            "StartDate": datetime.datetime(2026, 10, 1),
            "EndDate": datetime.datetime(2026, 10, 10),
        }
        values.update(overrides)
        return values

    def test_change_log_table_created_like_main_program(self):
        self.assertFalse(self.db.table_exists("plan_change_logs"))
        record_id = self.repos.plans.insert(self.plan_values(), self.OPERATOR)
        self.assertTrue(record_id > 0)
        self.assertTrue(self.db.table_exists("plan_change_logs"))
        for column in repositories.schema_generated.columns_of("plan_change_logs"):
            self.assertTrue(self.db.has_column("plan_change_logs", column), column)

    def test_insert_writes_row_and_add_log(self):
        record_id = self.repos.plans.insert(self.plan_values(), self.OPERATOR)
        row = self.repos.plans.get(record_id)
        self.assertEqual("31ABCDEFG7H", row["ModelName"])
        self.assertEqual("热冲击", row["TestItem"])
        # 日期按主程序格式落库（TEXT，无微秒）
        self.assertEqual("2026-10-01 00:00:00", row["StartDate"])
        self.assertEqual("2026-10-10 00:00:00", row["EndDate"])
        # 审计字段
        self.assertEqual(self.OPERATOR, row["CreatedBy"])
        self.assertEqual(self.OPERATOR, row["UpdatedBy"])
        self.assertTrue(row["CreatedAt"].startswith("20"))

        logs = self.repos.change_logs.recent()
        self.assertEqual(1, len(logs))
        self.assertEqual(repositories.ACTION_ADD, logs[0]["Action"])
        self.assertEqual("新增计划 RT2601-001 (31ABCDEFG7H)", logs[0]["Summary"])
        self.assertEqual(self.OPERATOR, logs[0]["Operator"])

        detail = self.db.query_one("SELECT * FROM plan_change_logs WHERE Id = ?", (logs[0]["Id"],))
        self.assertIsNone(detail["BeforeJson"])
        # PascalCase + 紧凑 JSON + ISO 日期（对齐 Newtonsoft）
        self.assertTrue(detail["AfterJson"].startswith('{"Id":'))
        self.assertIn('"ModelName":"31ABCDEFG7H"', detail["AfterJson"])
        self.assertIn('"StartDate":"2026-10-01T00:00:00"', detail["AfterJson"])
        self.assertNotIn(": ", detail["AfterJson"])

    def test_update_writes_before_after_and_skips_noop(self):
        record_id = self.repos.plans.insert(self.plan_values(), self.OPERATOR)
        self.assertTrue(self.repos.plans.update(record_id, {"Status": "Close", "Remark": "已结案"}, self.OPERATOR))
        row = self.repos.plans.get(record_id)
        self.assertEqual("Close", row["Status"])
        self.assertEqual("已结案", row["Remark"])

        logs = self.repos.change_logs.recent()
        self.assertEqual(repositories.ACTION_EDIT, logs[0]["Action"])
        self.assertEqual("编辑计划 RT2601-001 (31ABCDEFG7H)", logs[0]["Summary"])
        detail = self.db.query_one("SELECT * FROM plan_change_logs WHERE Id = ?", (logs[0]["Id"],))
        self.assertIn('"Status":"进行中"', detail["BeforeJson"])
        self.assertIn('"Status":"Close"', detail["AfterJson"])

        # 内容没变 → 不写库也不写日志
        before_count = self.repos.change_logs.count()
        self.assertFalse(self.repos.plans.update(record_id, {"Status": "Close"}, self.OPERATOR))
        self.assertEqual(before_count, self.repos.change_logs.count())

    def test_delete_writes_log_and_removes_row(self):
        record_id = self.repos.plans.insert(self.plan_values(), self.OPERATOR)
        self.assertTrue(self.repos.plans.delete(record_id, self.OPERATOR))
        self.assertIsNone(self.repos.plans.get(record_id))
        logs = self.repos.change_logs.recent()
        self.assertEqual(repositories.ACTION_DELETE, logs[0]["Action"])
        self.assertEqual("删除计划 RT2601-001 (31ABCDEFG7H)", logs[0]["Summary"])
        detail = self.db.query_one("SELECT * FROM plan_change_logs WHERE Id = ?", (logs[0]["Id"],))
        self.assertIsNotNone(detail["BeforeJson"])
        self.assertIsNone(detail["AfterJson"])

    def test_requisition_roundtrip_and_summary(self):
        record_id = self.repos.requisitions.insert(
            {
                "RequisitionDate": datetime.datetime(2026, 8, 12),
                "RequisitionNo": "2608-001",
                "ModelName": "31ABCDEFG7H",
                "OutQty": "2",
                "Disposition": "入库",
                "WorkOrder": "WO-1",
                "Remark": "含中文备注",
            },
            "张伟",
        )
        row = self.repos.requisitions.get(record_id)
        self.assertEqual("2608-001", row["RequisitionNo"])
        self.assertEqual("含中文备注", row["Remark"])
        self.assertEqual("2026-08-12 00:00:00", row["RequisitionDate"])
        self.assertEqual("新增领退 2608-001 (31ABCDEFG7H)", self.repos.change_logs.recent()[0]["Summary"])
        self.assertTrue(self.repos.requisitions.requisition_no_exists("2608-001"))
        self.assertFalse(self.repos.requisitions.requisition_no_exists("2608-001", exclude_id=record_id))

    def test_change_log_is_shared_by_plans_and_requisitions(self):
        """主程序的计划与领退共用 plan_change_logs，PlanId 会重号 —— 靠 Action/Summary 区分。"""
        plan_id = self.repos.plans.insert(self.plan_values(), self.OPERATOR)
        requisition_id = self.repos.requisitions.insert(
            {"RequisitionNo": "RT2601-001", "ModelName": "31ABCDEFG7H"}, self.OPERATOR
        )
        self.assertEqual(2, self.repos.change_logs.count())
        summaries = [row["Summary"] for row in self.repos.change_logs.recent()]
        self.assertTrue(any(item.startswith("新增计划") for item in summaries), summaries)
        self.assertTrue(any(item.startswith("新增领退") for item in summaries), summaries)
        if plan_id == requisition_id:
            rows = self.db.query('SELECT "Summary" FROM "plan_change_logs" WHERE "PlanId" = ?', (plan_id,))
            self.assertEqual(2, len(rows), "Id 重号时两条日志会同时命中 PlanId，必须靠 Summary 区分")

    def test_unique_index_still_enforced(self):
        self.repos.plans.insert(self.plan_values(), self.OPERATOR)
        with self.assertRaises(sqlite3.IntegrityError):
            self.repos.plans.insert(self.plan_values(), self.OPERATOR)
        self.assertTrue(self.repos.plans.job_no_exists("RT2601-001"))

    def test_search_and_count(self):
        self.repos.plans.insert(self.plan_values(), self.OPERATOR)
        self.repos.plans.insert(
            self.plan_values(JobNo="RT2601-002", ModelName="32XYZ", TestItem="Burn-In"), self.OPERATOR
        )
        self.assertEqual(2, self.repos.plans.count())
        self.assertEqual(1, len(self.repos.plans.list(keyword="XYZ")))
        self.assertEqual(2, len(self.repos.plans.list()))
        self.assertEqual(1, len(self.repos.plans.list(keyword="热冲击")))
        self.assertEqual(1, len(self.repos.plans.list(keyword="Burn-In")))

    def test_lookup_code_rules_match_main_program(self):
        lookups = self.repos.lookups
        self.assertEqual("31", lookups.model_to_product_code("31ABCDEFG7H"))
        self.assertEqual("FG", lookups.model_to_customer_code("31ABCDEFG7H"))
        # 主程序规则：客户代码取第 8-9 位，不足 9 位就没有
        self.assertIsNone(lookups.model_to_customer_code("SHORT"))
        self.assertEqual("RT", lookups.model_to_customer_code("TOO-SHORT"))
        self.assertIsNone(lookups.model_to_product_code(""))
        self.assertIsNone(lookups.model_to_customer_code(None))

    def test_lookup_queries(self):
        self.db.execute("INSERT INTO stages (Name, Description) VALUES (?, ?)", ("DVT", "设计验证"))
        self.db.execute("INSERT INTO stages (Name, Description) VALUES (?, ?)", ("MP", "量产"))
        self.assertEqual(["DVT", "MP"], self.repos.lookups.stages())

        self.db.execute(
            "INSERT INTO code_mappings (CodeType, Code, Name) VALUES (?, ?, ?)", ("P", "31", "电源供应器")
        )
        self.db.execute(
            "INSERT INTO code_mappings (CodeType, Code, Name) VALUES (?, ?, ?)", ("C", "FG", "客户甲")
        )
        self.assertEqual("电源供应器", self.repos.lookups.code_mapping("P", "31"))
        self.assertEqual("电源供应器", self.repos.lookups.find_product_by_model("31ABCDEFG7H"))
        self.assertEqual("客户甲", self.repos.lookups.find_customer_by_model("31ABCDEFG7H"))
        self.assertIsNone(self.repos.lookups.code_mapping("P", "99"))

        self.db.execute(
            "INSERT INTO model_mappings (ModelName, Product, Customer) VALUES (?, ?, ?)",
            ("31ABCDEFG7H", "电源供应器", "客户甲"),
        )
        mapping = self.repos.lookups.find_model_mapping(" 31ABCDEFG7H ")
        self.assertEqual("客户甲", mapping["Customer"])
        self.assertIsNone(self.repos.lookups.find_model_mapping("不存在"))
        self.assertEqual(1, len(self.repos.lookups.model_mappings()))

        self.db.execute("INSERT INTO test_items_catalog (Name) VALUES (?)", ("热冲击",))
        self.db.execute("INSERT INTO test_items_catalog (Name) VALUES (?)", ("Burn-In",))
        self.assertEqual(["Burn-In", "热冲击"], self.repos.lookups.test_items())

        self.db.execute("INSERT INTO products (Name, Code) VALUES (?, ?)", ("电源供应器", "31"))
        self.assertEqual(["电源供应器"], self.repos.lookups.products())
        self.db.execute("INSERT INTO customers (Name, Code) VALUES (?, ?)", ("客户甲", "FG"))
        self.assertEqual(["客户甲"], self.repos.lookups.customers())


class PlanRulesTests(unittest.TestCase):
    """校验规则与自动编号：逐条对齐主程序 PlanValidation / PlanExcelService。"""

    def test_job_no_format(self):
        rules = plan_rules
        for value in ("RT260801", "QRT260812", "rt260801", "RT2608999"):
            self.assertIsNone(rules.validate_job_no(value), value)
        for value in ("RT2608", "260801", "RT2608AB", "X260801", "RT123"):
            self.assertEqual(rules.MESSAGE_JOB_NO_FORMAT, rules.validate_job_no(value), value)
        self.assertIsNone(rules.validate_job_no(""))
        self.assertIsNone(rules.validate_job_no(None))
        self.assertEqual(rules.MESSAGE_JOB_NO_SEQ, rules.validate_job_no("RT260800"))

    def test_return_rt_order_format(self):
        rules = plan_rules
        self.assertIsNone(rules.validate_return_rt_order("RTAH260901"))
        self.assertIsNone(rules.validate_return_rt_order(None))
        self.assertEqual(rules.MESSAGE_RETURN_RT_FORMAT, rules.validate_return_rt_order("RTAH2609"))
        self.assertEqual(rules.MESSAGE_RETURN_RT_SEQ, rules.validate_return_rt_order("RTAH260900"))

    def test_status_validation_and_kind(self):
        rules = plan_rules
        self.assertIsNone(rules.validate_status("Close"))
        self.assertIsNone(rules.validate_status("ongoing"))
        self.assertEqual(rules.MESSAGE_STATUS, rules.validate_status("结案"))
        self.assertEqual(rules.STATUS_ONGOING, rules.status_kind("进行中"))
        self.assertEqual(rules.STATUS_ONGOING, rules.status_kind("Ongoing"))
        self.assertEqual(rules.STATUS_PENDING, rules.status_kind("待测"))
        self.assertEqual(rules.STATUS_CLOSED, rules.status_kind("Close"))
        self.assertEqual(rules.STATUS_CLOSED, rules.status_kind("已完成"))
        self.assertEqual("", rules.status_kind("随便写的"))
        self.assertEqual("", rules.status_kind(None))

    def test_catalog_and_disposition_validation(self):
        rules = plan_rules
        self.assertIsNone(rules.validate_in_catalog("DVT", ["DVT", "MP"], "阶段"))
        self.assertIsNone(rules.validate_in_catalog("", ["DVT"], "阶段"))
        self.assertIn("不在字典中", rules.validate_in_catalog("EVT", ["DVT"], "阶段"))
        self.assertIsNone(rules.validate_disposition("入库"))
        self.assertIsNone(rules.validate_disposition("报废"))
        self.assertIn("只能是", rules.validate_disposition("丢弃"))

    def test_sequence_formatting(self):
        self.assertEqual("01", plan_rules.format_sequence(1))
        self.assertEqual("99", plan_rules.format_sequence(99))
        self.assertEqual("100", plan_rules.format_sequence(100))

    def test_generate_job_no_shares_monthly_sequence(self):
        existing = ["RT260801", "QRT260802", "QRT261001", "RT259912", None, "垃圾数据"]
        self.assertEqual("RT260803", plan_rules.generate_job_no(existing, datetime.datetime(2026, 8, 15)))
        self.assertEqual("QRT260803", plan_rules.generate_job_no(existing, datetime.datetime(2026, 8, 15), "QRT"))
        self.assertEqual("RT260901", plan_rules.generate_job_no(existing, datetime.datetime(2026, 9, 1)))
        self.assertEqual("QRT261002", plan_rules.generate_job_no(existing, datetime.datetime(2026, 10, 1), "QRT"))
        # 序号 ≥ 100 时按实际位数展开
        self.assertEqual("RT2608100", plan_rules.generate_job_no(["RT260899"], datetime.datetime(2026, 8, 1)))

    def test_generate_return_rt_order(self):
        existing = ["RTAH260901", "RTAH260902", "RTAH261001", "无效"]
        self.assertEqual("RTAH260903", plan_rules.generate_return_rt_order(existing, datetime.datetime(2026, 9, 20)))
        self.assertEqual("RTAH261002", plan_rules.generate_return_rt_order(existing, datetime.datetime(2026, 10, 2)))


class EditServiceTests(TempDatabaseTest):
    """编辑服务：必填/格式/唯一性校验 + 落库 + 变更日志（不依赖界面）。"""

    OPERATOR = "tester"

    def setUp(self):
        TempDatabaseTest.setUp(self)
        self.repos = repositories.Repositories(self.db)
        self.requisitions = plans_service.RequisitionService(self.repos, self.repos.lookups, self.OPERATOR)
        self.plans = plans_service.PlanService(self.repos, self.repos.lookups, self.OPERATOR)
        self.db.execute("INSERT INTO stages (Name) VALUES (?)", ("DVT",))
        self.db.execute("INSERT INTO test_items_catalog (Name) VALUES (?)", ("热冲击",))

    def requisition_values(self, **overrides):
        values = {
            "RequisitionDate": datetime.datetime(2026, 9, 1),
            "RequisitionNo": "2609-001",
            "ModelName": "31ABCDEFG7H",
            "OutQty": "2",
            "Rev": "A",
            "WorkOrder": "WO-1",
            "Disposition": plan_rules.DISPOSITION_STOCK_IN,
            "ReturnRtOrder": "RTAH260901",
            "Remark": "备注",
        }
        values.update(overrides)
        return values

    def plan_values(self, **overrides):
        values = {
            "ModelName": "31ABCDEFG7H",
            "TestItem": "热冲击",
            "Stage": "DVT",
            "Owner": "王工",
            "StartDate": datetime.datetime(2026, 9, 1),
            "Status": "Ongoing",
            "Remark": "计划备注",
        }
        values.update(overrides)
        return values

    def test_requisition_required_fields(self):
        result = self.requisitions.save({})
        self.assertFalse(result.ok)
        self.assertIn("请选择领用日期", result.errors)
        self.assertIn("请填写領料單据號", result.errors)
        self.assertIn("请填写机种名", result.errors)
        self.assertIn("请填写领用数量", result.errors)
        self.assertIn("请填写版本", result.errors)
        self.assertIn("请填写工令", result.errors)
        self.assertEqual(0, self.repos.requisitions.count())

    def test_requisition_disposition_and_return_rt_rules(self):
        result = self.requisitions.save(
            self.requisition_values(Disposition=plan_rules.DISPOSITION_STOCK_IN, ReturnRtOrder="")
        )
        self.assertFalse(result.ok)
        self.assertIn("单体去向为「入库」时必须填写回线RT工令", result.errors)

        result = self.requisitions.save(self.requisition_values(ReturnRtOrder="RT260901"))
        self.assertFalse(result.ok)
        self.assertIn(plan_rules.MESSAGE_RETURN_RT_FORMAT, result.errors)

        # 报废无需回线RT工令
        result = self.requisitions.save(
            self.requisition_values(Disposition=plan_rules.DISPOSITION_SCRAP, ReturnRtOrder="")
        )
        self.assertTrue(result.ok, result.errors)

    def test_requisition_unique_rules(self):
        first = self.requisitions.save(self.requisition_values())
        self.assertTrue(first.ok, first.errors)
        duplicate_no = self.requisitions.save(self.requisition_values(ReturnRtOrder="RTAH260902"))
        self.assertFalse(duplicate_no.ok)
        self.assertIn("領料單据號 [2609-001] 已存在", duplicate_no.errors)
        duplicate_rt = self.requisitions.save(self.requisition_values(RequisitionNo="2609-002"))
        self.assertFalse(duplicate_rt.ok)
        self.assertIn("回线RT工令 [RTAH260901] 已存在", duplicate_rt.errors)
        # 编辑自己时不算重号
        result = self.requisitions.save(self.requisition_values(), first.record_id)
        self.assertTrue(result.ok, result.errors)

    def test_requisition_save_edit_delete_writes_logs(self):
        created = self.requisitions.save(self.requisition_values())
        self.assertTrue(created.ok, created.errors)
        row = self.repos.requisitions.get(created.record_id)
        self.assertEqual("2026-09-01 00:00:00", row["RequisitionDate"])
        self.assertEqual(self.OPERATOR, row["CreatedBy"])

        edited = self.requisitions.save(self.requisition_values(OutQty="5"), created.record_id)
        self.assertTrue(edited.ok, edited.errors)
        self.assertEqual("5", self.repos.requisitions.get(created.record_id)["OutQty"])

        deleted = self.requisitions.delete(created.record_id)
        self.assertTrue(deleted.ok)
        self.assertIsNone(self.repos.requisitions.get(created.record_id))

        actions = [item["Action"] for item in self.repos.change_logs.recent()]
        self.assertEqual([repositories.ACTION_DELETE, repositories.ACTION_EDIT, repositories.ACTION_ADD], actions)
        summaries = [item["Summary"] for item in self.repos.change_logs.recent()]
        self.assertEqual("删除领退 2609-001 (31ABCDEFG7H)", summaries[0])

    def test_requisition_return_rt_generator(self):
        self.assertEqual("RTAH260901", self.requisitions.next_return_rt_order(datetime.datetime(2026, 9, 5)))
        self.requisitions.save(self.requisition_values())
        self.assertEqual("RTAH260902", self.requisitions.next_return_rt_order(datetime.datetime(2026, 9, 6)))

    def test_plan_required_fields_and_catalog(self):
        result = self.plans.save({})
        self.assertFalse(result.ok)
        self.assertIn("请选择测试项目", result.errors)
        self.assertIn("请填写开始时间", result.errors)
        self.assertIn("请选择阶段", result.errors)
        self.assertIn("请填写机种名", result.errors)
        self.assertIn("请填写备注", result.errors)

        result = self.plans.save(self.plan_values(TestItem="不存在的项目"))
        self.assertFalse(result.ok)
        self.assertIn("不在字典中", "\n".join(result.errors))

        result = self.plans.save(self.plan_values(Status="结案"))
        self.assertFalse(result.ok)
        self.assertIn(plan_rules.MESSAGE_STATUS, result.errors)

    def test_plan_auto_job_no_and_uniqueness(self):
        first = self.plans.save(self.plan_values())
        self.assertTrue(first.ok, first.errors)
        first_row = self.repos.plans.get(first.record_id)
        self.assertEqual("QRT260901", first_row["JobNo"])

        second = self.plans.save(self.plan_values(ModelName="31OTHER"))
        self.assertEqual("QRT260902", self.repos.plans.get(second.record_id)["JobNo"])

        duplicate = self.plans.save(self.plan_values(JobNo="QRT260901"))
        self.assertFalse(duplicate.ok)
        self.assertIn("工作编号 [QRT260901] 已存在", duplicate.errors)

        edited = self.plans.save(self.plan_values(JobNo="QRT260901", Status="Close"), first.record_id)
        self.assertTrue(edited.ok, edited.errors)
        self.assertEqual("Close", self.repos.plans.get(first.record_id)["Status"])

    def test_plan_suggest_product_customer(self):
        self.db.execute(
            "INSERT INTO model_mappings (ModelName, Product, Customer) VALUES (?, ?, ?)",
            ("31ABCDEFG7H", "电源供应器", "客户甲"),
        )
        product, customer = self.plans.suggest_product_customer("31ABCDEFG7H")
        self.assertEqual("电源供应器", product)
        self.assertEqual("客户甲", customer)
        self.assertEqual((None, None), self.plans.suggest_product_customer(None))


class FormHelperTests(unittest.TestCase):
    """表单字段与「界面字符串 ↔ 服务值」转换（放在服务层，可脱离界面测试）。"""

    class FakeLookups(object):
        def test_items(self):
            return ["热冲击"]

        def stages(self):
            return ["DVT"]

        def products(self):
            return ["电源供应器"]

        def customers(self):
            return ["客户甲"]

    def test_form_fields_cover_payload_fields(self):
        fields = set(key for key, _label, _kind, _options in plans_service.requisition_form_fields())
        payload = set(plans_service.REQUISITION_TEXT_FIELDS) | set(plans_service.REQUISITION_DATE_FIELDS)
        missing = payload - fields - set(plans_service.FORM_EXEMPT_FIELDS)
        self.assertEqual(set(), missing, "领退表单缺字段：%s" % missing)

        plan_fields = set(
            key for key, _label, _kind, _options in plans_service.plan_form_fields(self.FakeLookups())
        )
        plan_payload = set(plans_service.PLAN_TEXT_FIELDS) | set(plans_service.PLAN_DATE_FIELDS)
        self.assertEqual(set(), plan_payload - plan_fields, "计划表单缺字段")

    def test_plan_form_fields_use_catalogs(self):
        fields = dict(
            (key, options) for key, _label, _kind, options in plans_service.plan_form_fields(self.FakeLookups())
        )
        self.assertEqual(("热冲击",), fields["TestItem"])
        self.assertEqual(("DVT",), fields["Stage"])
        self.assertEqual(("电源供应器",), fields["Product"])
        self.assertEqual(("客户甲",), fields["Customer"])
        self.assertEqual(plan_rules.VALID_STATUSES, fields["Status"])

    def test_form_values_conversion(self):
        values = plans_service.form_values(
            {"RequisitionDate": "2026-10-07", "OutQty": "", "Remark": "", "RequisitionNo": "2609-001"},
            plans_service.REQUISITION_DATE_FIELDS,
            ("Remark",),
        )
        self.assertEqual(datetime.datetime(2026, 10, 7), values["RequisitionDate"])
        self.assertIsNone(values["OutQty"])
        self.assertEqual("", values["Remark"])
        self.assertEqual("2609-001", values["RequisitionNo"])

        values = plans_service.form_values(
            {"StartDate": "不是日期", "EndDate": ""}, plans_service.PLAN_DATE_FIELDS
        )
        self.assertIsNone(values["StartDate"])
        self.assertIsNone(values["EndDate"])

    def test_form_initial_conversion(self):
        fields = plans_service.requisition_form_fields()
        row = dict((key, None) for key, _label, _kind, _options in fields)
        row["RequisitionDate"] = "2026-10-07 00:00:00"
        row["RequisitionNo"] = "2609-001"
        initial = plans_service.form_initial(fields, row, plans_service.REQUISITION_DATE_FIELDS)
        self.assertEqual("2026-10-07", initial["RequisitionDate"])
        self.assertEqual("2609-001", initial["RequisitionNo"])
        self.assertEqual("", initial["Remark"])


class StartupFailureTests(unittest.TestCase):
    """启动期失败必须「看得见」。

    打包后的 exe 是窗口子系统：没有控制台、``sys.stderr`` 是 ``None``，往 stderr 写东西
    等于没写。踩过的坑就是「双击 exe 只写了两行日志就没反应」——数据目录缺失时 sqlite 会
    顺手建一个空库，或者直接退出且不解释。这里把这些路径全部钉住。
    """

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp(prefix="ort-xp-startup-")
        self.context = context_module.AppContext(
            app_directory=self.temp_dir,
            data_folder=self.temp_dir,
            logger=logging.getLogger("ort_xp.startup-test"),
        )
        # 测试进程里绝不弹消息框（否则会卡住），直接跑完再恢复
        fatal.set_dialogs_enabled(False)

    def tearDown(self):
        fatal.set_dialogs_enabled(True)
        logger = logging.getLogger(logging_setup.LOGGER_NAME)
        for handler in list(logger.handlers):
            logger.removeHandler(handler)
            try:
                handler.close()
            except Exception:
                pass
        shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_missing_database_text_names_every_way_out(self):
        text = main_module.missing_database_text(self.context)
        self.assertIn(self.context.db_path, text)
        self.assertIn(self.context.local_settings_path, text)
        self.assertIn("ort_plans.db", text)
        self.assertIn("--data-folder", text)

    def test_open_failure_text_keeps_the_reason(self):
        text = main_module.open_failure_text(self.context, RuntimeError("磁盘满了"))
        self.assertIn("磁盘满了", text)
        self.assertIn(self.context.db_path, text)

    def test_report_writes_the_log_without_a_logger(self):
        self.assertEqual(2, fatal.report("没有找到数据库", "正文：DATA", self.temp_dir))
        text = compat.read_text(os.path.join(self.temp_dir, "Logs", "ort_xp.log"))
        self.assertIn("[ERROR]", text)
        self.assertIn("没有找到数据库：正文：DATA", text)

    def test_report_never_raises_when_the_log_is_unwritable(self):
        self.assertEqual(2, fatal.report("标题", "正文", os.path.join(self.temp_dir, "no", "such", "dir")))

    def test_dialog_switch_follows_the_env_var(self):
        os.environ[fatal.NO_DIALOG_ENV] = "1"
        try:
            self.assertFalse(fatal.dialogs_enabled())
        finally:
            os.environ.pop(fatal.NO_DIALOG_ENV, None)
        fatal.set_dialogs_enabled(True)
        self.assertTrue(fatal.dialogs_enabled())

    def test_first_run_finds_database_in_folder_or_its_data_subdir(self):
        self.assertIsNone(first_run.find_database_folder(self.temp_dir))
        self.assertIsNone(first_run.find_database_folder(""))
        empty = os.path.join(self.temp_dir, "empty")
        compat.ensure_dir(empty)
        self.assertIsNone(first_run.find_database_folder(empty))

        data_dir = os.path.join(self.temp_dir, "Data")
        compat.ensure_dir(data_dir)
        compat.write_text(os.path.join(data_dir, "ort_plans.db"), "")
        # 用户选到主程序目录（bin\Debug）时，要能自己往下找一层 Data
        self.assertEqual(data_dir, first_run.find_database_folder(self.temp_dir))
        self.assertEqual(data_dir, first_run.find_database_folder(data_dir))

        compat.write_text(os.path.join(self.temp_dir, "ort_plans.db"), "")
        self.assertEqual(self.temp_dir, first_run.find_database_folder(self.temp_dir))

    def test_no_picker_when_the_folder_was_given_or_dialogs_are_off(self):
        explicit = argparse.Namespace(data_folder="Z:\\ORTData")
        self.assertIsNone(main_module.choose_data_folder(self.context, explicit, "原因"))
        implicit = argparse.Namespace(data_folder=None)
        self.assertIsNone(main_module.choose_data_folder(self.context, implicit, "原因"))

    def test_headless_main_reports_missing_database_and_creates_nothing(self):
        missing = os.path.join(self.temp_dir, "nope")
        code = main_module.main(
            ["--app-dir", self.temp_dir, "--data-folder", missing, "--no-dialog"]
        )
        self.assertEqual(2, code)
        # 关键：不能因为「连得上一个空库」就把程序放进去，也不能凭空造一个库出来
        self.assertFalse(os.path.isfile(os.path.join(missing, "ort_plans.db")))
        text = compat.read_text(os.path.join(self.temp_dir, "Logs", "ort_xp.log"))
        self.assertIn("没有找到数据库", text)


class SourceCompatibilityTests(unittest.TestCase):
    def test_package_is_python34_compatible(self):
        findings = compat_check.check_path(os.path.join(_PARENT, "ort_xp"))
        self.assertEqual([], findings)


if __name__ == "__main__":
    unittest.main()
