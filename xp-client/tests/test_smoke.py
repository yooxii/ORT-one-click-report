# -*- coding: utf-8 -*-
"""标准库 unittest（XP 上不引入第三方测试框架）。

运行::

    python -m unittest discover -s tests -v
"""

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
from ort_xp import compat, config, credentials, dpapi  # noqa: E402
from ort_xp.db.connection import Database  # noqa: E402
from ort_xp.services import auth as auth_service  # noqa: E402
from ort_xp.services import mail as mail_service  # noqa: E402

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
    WorkOrder NVARCHAR(64) NULL,
    ReturnRtOrder NVARCHAR(64) NULL,
    Remark NVARCHAR(512) NULL,
    CreatedAt DATETIME NULL
);
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


class DatabaseTests(TempDatabaseTest):
    def test_table_and_column_detection(self):
        self.assertTrue(self.db.table_exists("users"))
        self.assertFalse(self.db.table_exists("plan_change_logs"))
        self.assertIn("Username", self.db.column_names("users"))
        self.assertTrue(self.db.has_column("users", "username"))
        self.assertFalse(self.db.has_column("users", "NotThere"))

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


class SourceCompatibilityTests(unittest.TestCase):
    def test_package_is_python34_compatible(self):
        findings = compat_check.check_path(os.path.join(_PARENT, "ort_xp"))
        self.assertEqual([], findings)


if __name__ == "__main__":
    unittest.main()
