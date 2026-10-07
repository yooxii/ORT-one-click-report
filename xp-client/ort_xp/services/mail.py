# -*- coding: utf-8 -*-
"""邮件：SMTP 发送、收件人解析、发送记录与去重、计划到期提醒。

对应主程序 ``Services/MailService.cs``、``SmtpMailClient.cs``、``MailNotifier.cs``。

XP 上的关键差异：Python 自带 OpenSSL，**不受系统 Schannel 只支持 TLS 1.0 的限制**，
所以 TLS 1.2 的 STARTTLS / 隐式 SSL 都能用（这正是选 Python 而不是 .NET 4.0 的原因）。

不支持：``mail.useDefaultCredentials``（Windows 集成验证），Python 标准库没有 NTLM。
"""

import datetime
import os
import re
import smtplib
import ssl
from email.header import Header
from email.mime.text import MIMEText
from email.utils import formataddr

from .. import compat
from ..db.repositories import PlanRepository

MAIL_KIND_NOTICE = "Notice"
MAIL_KIND_WARNING = "Warning"

#: 主程序未自定义模板时使用的内置默认（与主程序资源里的默认模板同义）
DEFAULT_SUBJECTS = {
    MAIL_KIND_NOTICE: "{{Title}} - ORT实验室管理系统",
    MAIL_KIND_WARNING: "{{JobNo}} 测试结束日期提醒（{{EndDate|yyyy/M/d}}）",
}

DEFAULT_BODIES = {
    MAIL_KIND_NOTICE: "{{DisplayName}} 您好：\n\n{{Content}}\n\n类型：{{Category}}\n时间：{{Time}}",
    MAIL_KIND_WARNING: (
        "{{Owner}} 您好：\n\n"
        "以下计划的测试结束日期临近或已逾期，且完成状况仍未结案，请及时处理。\n\n"
        "工作編號：{{JobNo}}\n机种：{{ModelName}}\n结束日期：{{EndDate|yyyy/M/d}}\n"
        "剩余天数：{{DaysLeft}}\n完成状况：{{Status}}"
    ),
}

_ADDRESS_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
_TEMPLATE_RE = re.compile(r"\{\{\s*([A-Za-z0-9_]+)\s*(?:\|\s*([^}]+?)\s*)?\}\}")


def is_valid_address(address):
    return bool(address) and bool(_ADDRESS_RE.match(address.strip()))


def render_template(text, variables):
    """渲染主程序模板语法：``{{Name}}`` 与 ``{{Name|yyyy/M/d}}``。"""

    def _replace(match):
        name = match.group(1)
        pattern = match.group(2)
        value = variables.get(name)
        if value is None:
            return ""
        if pattern:
            return compat.format_dotnet_date(value, pattern)
        return str(value)

    return _TEMPLATE_RE.sub(_replace, text or "")


def split_list(text):
    """拆分分号/逗号/顿号分隔的地址列表。"""
    if not text:
        return []
    normalized = text
    for separator in (";", ",", "，", "、"):  # 分号、半角逗号、全角逗号、顿号
        normalized = normalized.replace(separator, ";")
    return [item.strip() for item in normalized.split(";") if item.strip()]


def is_completed(status):
    """主程序规则：完成状况为 Close（大小写无关，空值视为未完成）。"""
    return (status or "").strip().lower() == "close"


class MailResult(object):
    def __init__(self, success=False, skipped=False, message="", subject="", recipients=None):
        self.success = success
        self.skipped = skipped
        self.message = message
        self.subject = subject
        self.recipients = recipients or []

    def __repr__(self):
        state = "成功" if self.success else ("跳过" if self.skipped else "失败")
        return "<MailResult %s %s>" % (state, self.message)


class MailService(object):
    def __init__(self, database, settings, credential_store=None, logger=None, ca_bundle=None):
        self._db = database
        self.settings = settings
        self._credentials = credential_store
        self._logger = logger
        self.ca_bundle = ca_bundle

    # ---------------------------------------------------------------- 可用性

    def is_ready(self):
        """返回 ``(是否可发, 原因)``。"""
        settings = self.settings
        if not settings.enabled:
            return False, "邮件功能未启用（mail.enabled）"
        if not settings.host:
            return False, "未配置 SMTP 服务器（mail.host）"
        if not is_valid_address(settings.from_address):
            return False, "发件地址无效或为空（mail.fromAddress）"
        if settings.use_default_credentials:
            return False, "不支持 Windows 集成验证（mail.useDefaultCredentials），请改用账号密码"
        return True, ""

    # ---------------------------------------------------------------- 发送

    def send(self, kind, recipients, subject, body, ref_type=None, ref_key=None, dry_run=False):
        """发送一封邮件；无论成功失败都写 ``mail_logs``。"""
        ready, reason = self.is_ready()
        if not ready:
            return MailResult(skipped=True, message=reason)

        recipients = self.normalize_recipients(recipients)
        if not recipients:
            return MailResult(skipped=True, message="没有有效的收件人地址")

        if ref_type and ref_key:
            days = self.settings.dedupe_days
            if days > 0 and self.was_sent_recently(kind, ref_type, ref_key, days):
                return MailResult(
                    skipped=True,
                    message="去重：%s 天内已对 %s/%s 发过同类邮件" % (days, ref_type, ref_key),
                    subject=subject,
                    recipients=recipients,
                )

        if dry_run:
            return MailResult(
                success=True,
                message="dry-run：未真正发送",
                subject=subject,
                recipients=recipients,
            )

        error = ""
        try:
            self._smtp_send(recipients, subject, body)
            success = True
        except Exception as exc:
            success = False
            error = "%s: %s" % (exc.__class__.__name__, exc)
            if self._logger is not None:
                self._logger.error("邮件发送失败：%s" % error)

        self._write_log(kind, ref_type, ref_key, recipients, subject, success, error)
        if success:
            return MailResult(success=True, message="已发送", subject=subject, recipients=recipients)
        return MailResult(success=False, message=error, subject=subject, recipients=recipients)

    def send_test(self, recipient, dry_run=False):
        subject = "ORT实验室管理系统 XP 客户端 · 测试邮件"
        body = "这是一封测试邮件，用于验证 SMTP 配置。\n\n发送时间：%s" % compat.format_datetime_text(
            datetime.datetime.now()
        )
        return self.send(MAIL_KIND_NOTICE, [recipient], subject, body, dry_run=dry_run)

    def _smtp_send(self, recipients, subject, body):
        settings = self.settings
        message = MIMEText(body or "", "html" if settings.body_is_html else "plain", "utf-8")
        message["Subject"] = Header(subject or "", "utf-8")
        message["From"] = formataddr((str(Header(settings.from_name or "", "utf-8")), settings.from_address))
        message["To"] = ", ".join(recipients)

        timeout = max(5, settings.timeout_seconds or 30)
        security = (settings.security or "None").strip().lower()

        if security == "ssl":
            server = smtplib.SMTP_SSL(settings.host, settings.port, timeout=timeout, context=self._ssl_context())
        else:
            server = smtplib.SMTP(settings.host, settings.port, timeout=timeout)
        try:
            try:
                server.ehlo()
            except smtplib.SMTPException:
                pass
            if security == "starttls":
                server.starttls(context=self._ssl_context())
                server.ehlo()
            if settings.username:
                server.login(settings.username, settings.password or "")
            server.sendmail(settings.from_address, recipients, message.as_string())
        finally:
            try:
                server.quit()
            except Exception:
                pass

    def _ssl_context(self):
        """构造 SSL 上下文。

        只用 Python 3.4 就有的 API（``ssl.TLSVersion`` 是 3.7 才有的，不能用）。
        Python 3.4 的 ``ssl`` 不会自动加载 Windows 证书库，而 XP 的根证书库早已过期，
        所以优先用随程序分发的 PEM；没有时回落到系统默认，并把错误说清楚。
        """
        settings = self.settings
        if settings.ignore_cert_errors:
            context = ssl.SSLContext(ssl.PROTOCOL_TLSv1_2)
            context.check_hostname = False
            context.verify_mode = ssl.CERT_NONE
            return context
        if self.ca_bundle and os.path.isfile(self.ca_bundle):
            return ssl.create_default_context(cafile=self.ca_bundle)
        try:
            return ssl.create_default_context()
        except Exception as exc:
            raise RuntimeError(
                "无法建立 TLS 上下文（多半是系统根证书不可用）：%s。"
                "请把一份 cacert.pem 放到程序目录（或 ca\\cacert.pem）后重试，"
                "或在设置里显式开启「忽略证书错误」（内网自签名服务器）。" % exc
            )

    # ---------------------------------------------------------------- 记录与去重

    def _write_log(self, kind, ref_type, ref_key, recipients, subject, success, error):
        if not self._db.table_exists("mail_logs"):
            if self._logger is not None:
                self._logger.warn("mail_logs 表不存在（由主程序创建），本次发送未记账")
            return
        try:
            self._db.execute(
                "INSERT INTO mail_logs (Kind, RefType, RefKey, Recipients, Subject, Success, Error, CreatedAt) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                (
                    kind,
                    ref_type,
                    ref_key,
                    ";".join(recipients)[:1024],
                    (subject or "")[:512],
                    1 if success else 0,
                    (error or "")[:1024],
                    compat.format_datetime_text(datetime.datetime.now()),
                ),
            )
        except Exception as exc:
            if self._logger is not None:
                self._logger.warn("写 mail_logs 失败：%s" % exc)

    def was_sent_recently(self, kind, ref_type, ref_key, days):
        if not self._db.table_exists("mail_logs"):
            return False
        since = compat.format_datetime_text(datetime.datetime.now() - datetime.timedelta(days=days))
        row = self._db.query_one(
            "SELECT Id FROM mail_logs WHERE Kind = ? AND RefType = ? AND RefKey = ? AND Success = 1 "
            "AND CreatedAt >= ? LIMIT 1",
            (kind, ref_type, ref_key, since),
        )
        return row is not None

    def recent_logs(self, count=20):
        if not self._db.table_exists("mail_logs"):
            return []
        return self._db.query(
            "SELECT Kind, RefType, RefKey, Recipients, Subject, Success, Error, CreatedAt "
            "FROM mail_logs ORDER BY Id DESC LIMIT ?",
            (count,),
        )

    # ---------------------------------------------------------------- 收件人解析

    def normalize_recipients(self, recipients):
        result = []
        seen = {}
        for item in recipients or []:
            address = (item or "").strip()
            if not is_valid_address(address):
                continue
            key = address.lower()
            if key in seen:
                continue
            seen[key] = True
            result.append(address)
        return result

    def resolve_user_email(self, name_or_username):
        """按显示名优先、登录名次之解析邮箱（与主程序一致，忽略大小写）。"""
        if not name_or_username:
            return None
        wanted = name_or_username.strip().lower()
        rows = self._db.query("SELECT Username, DisplayName, Email FROM users")
        fallback = None
        for row in rows:
            email = (row["Email"] or "").strip()
            if not is_valid_address(email):
                continue
            if (row["DisplayName"] or "").strip().lower() == wanted:
                return email
            if (row["Username"] or "").strip().lower() == wanted and fallback is None:
                fallback = email
        return fallback

    def role_emails(self, role):
        if not self._db.table_exists("user_roles"):
            return []
        rows = self._db.query(
            "SELECT u.Email FROM users u JOIN user_roles r ON r.UserId = u.Id "
            "WHERE lower(r.Role) = lower(?) AND u.IsActive = 1",
            (role,),
        )
        return self.normalize_recipients([row["Email"] for row in rows])


class PlanDeadlineReminder(object):
    """计划到期提醒（对应主程序 ``MailNotifier.CheckPlanDeadlines``）。"""

    REF_TYPE = "Plan"

    def __init__(self, database, mail_service, settings, logger=None, plans=None):
        self._db = database
        self._mail = mail_service
        self.settings = settings
        self._logger = logger
        # 计划查询统一走仓储，避免 SQL 散落在两处
        self._plans = plans if plans is not None else PlanRepository(database, logger=logger)

    def collect(self, today=None):
        """返回需要提醒的计划列表（已过滤已结案与不满足规则的项）。"""
        today = today or datetime.datetime.now()
        limit = today + datetime.timedelta(days=max(0, self.settings.warning_days_before))
        rows = self._plans.deadline_candidates(compat.format_datetime_text(limit))
        result = []
        for row in rows:
            if is_completed(row["Status"]):
                continue
            end_date = compat.parse_datetime_text(row["EndDate"])
            if end_date is None:
                continue
            days_left = (end_date.date() - today.date()).days
            if days_left < 0 and not self.settings.warning_include_overdue:
                continue
            result.append(
                {
                    "id": row["Id"],
                    "job_no": row["JobNo"] or "",
                    "model_name": row["ModelName"] or "",
                    "owner": row["Owner"] or "",
                    "end_date": end_date,
                    "days_left": days_left,
                    "overdue": days_left < 0,
                    "status": row["Status"] or "",
                }
            )
        return result

    def run(self, dry_run=False, today=None):
        """执行一轮提醒，返回统计信息。"""
        summary = {"candidates": 0, "sent": 0, "skipped": 0, "failed": 0, "details": []}
        if not self.settings.warning_enabled:
            summary["details"].append("警告类邮件未启用（mail.warningEnabled）")
            return summary

        plans = self.collect(today)
        summary["candidates"] = len(plans)
        for plan in plans:
            email = self._mail.resolve_user_email(plan["owner"])
            if not email:
                summary["skipped"] += 1
                summary["details"].append("%s：负责人「%s」没有可用邮箱" % (plan["job_no"], plan["owner"]))
                continue
            variables = {
                "JobNo": plan["job_no"],
                "ModelName": plan["model_name"],
                "Owner": plan["owner"],
                "EndDate": plan["end_date"],
                "DaysLeft": plan["days_left"],
                "Overdue": "是" if plan["overdue"] else "否",
                "Status": plan["status"],
                "AppName": "ORT实验室管理系统",
            }
            subject = render_template(
                self.settings.template(MAIL_KIND_WARNING, True) or DEFAULT_SUBJECTS[MAIL_KIND_WARNING], variables
            )
            body = render_template(
                self.settings.template(MAIL_KIND_WARNING, False) or DEFAULT_BODIES[MAIL_KIND_WARNING], variables
            )
            result = self._mail.send(
                MAIL_KIND_WARNING,
                [email],
                subject,
                body,
                ref_type=self.REF_TYPE,
                ref_key=plan["job_no"],
                dry_run=dry_run,
            )
            if result.success:
                summary["sent"] += 1
            elif result.skipped:
                summary["skipped"] += 1
            else:
                summary["failed"] += 1
            summary["details"].append("%s → %s：%s" % (plan["job_no"], email, result.message))
        return summary
