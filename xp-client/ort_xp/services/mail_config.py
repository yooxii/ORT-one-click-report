# -*- coding: utf-8 -*-
"""邮件设置的校验与写库（界面与测试共用，不依赖 tkinter）。

键名、取值范围、空值处理都与主程序 ``Services/AppSettingsService.cs`` 的 ``FillMailValues``
一一对应，写出来的文本格式（``True`` / ``False``、整数、空串写 NULL）也一致，
所以两端改的是同一份设置，互相都能读懂。

**刻意不写**的两类键：

- ``mail.passwordEnc``：那是主程序在**它那台机器**上用 DPAPI(CurrentUser) 加密的密文，
  XP 端写进去主程序也解不开（反之亦然）；XP 端的口令单独存在本机（见 ``credentials.py``）。
- ``mail.template.*`` / ``mail.ccAdmin.*``：模板与抄送管理员开关由主程序维护，XP 端只读。
"""

from . import mail as mail_service
from .. import compat

#: SMTP 安全方式（与主程序 ``MailSecurityOption`` 的 Code 一致）
SECURITY_OPTIONS = (
    ("None", "不加密"),
    ("StartTls", "STARTTLS"),
    ("Ssl", "SSL/TLS"),
)
SECURITY_CODES = tuple(code for code, _label in SECURITY_OPTIONS)

#: 界面字段定义：``kind`` 取 bool / text / int / choice；``attr`` 对应 MailSettings 的属性名
FIELDS = (
    {"key": "mail.enabled", "label": "启用邮件功能", "kind": "bool", "attr": "enabled"},
    {"key": "mail.noticeEnabled", "label": "发送通知类邮件", "kind": "bool", "attr": "notice_enabled"},
    {"key": "mail.warningEnabled", "label": "发送警告类邮件", "kind": "bool", "attr": "warning_enabled"},
    {"key": "mail.warningIncludeOverdue", "label": "提醒包含已逾期计划", "kind": "bool", "attr": "warning_include_overdue"},
    {"key": "mail.ignoreCertErrors", "label": "忽略证书错误", "kind": "bool", "attr": "ignore_cert_errors"},
    {"key": "mail.bodyIsHtml", "label": "正文按 HTML 发送", "kind": "bool", "attr": "body_is_html"},
    {
        "key": "mail.useDefaultCredentials",
        "label": "Windows 集成验证",
        "kind": "bool",
        "attr": "use_default_credentials",
    },
    {"key": "mail.host", "label": "SMTP 服务器", "kind": "text", "attr": "host"},
    {"key": "mail.port", "label": "端口", "kind": "int", "attr": "port", "min": 1, "max": 65535},
    {"key": "mail.security", "label": "安全方式", "kind": "choice", "attr": "security", "options": SECURITY_OPTIONS},
    {"key": "mail.fromAddress", "label": "发件地址", "kind": "text", "attr": "from_address"},
    {"key": "mail.fromName", "label": "发件人显示名", "kind": "text", "attr": "from_name"},
    {"key": "mail.username", "label": "SMTP 账号", "kind": "text", "attr": "username"},
    {"key": "mail.ccList", "label": "固定抄送", "kind": "text", "attr": "cc_list"},
    {"key": "mail.timeoutSeconds", "label": "超时（秒）", "kind": "int", "attr": "timeout_seconds", "min": 5, "max": 300},
    {
        "key": "mail.warningDaysBefore",
        "label": "提前提醒（天）",
        "kind": "int",
        "attr": "warning_days_before",
        "min": 0,
        "max": 365,
    },
    {"key": "mail.dedupeDays", "label": "去重天数", "kind": "int", "attr": "dedupe_days", "min": 0, "max": 365},
)

FIELD_KEYS = tuple(field["key"] for field in FIELDS)

#: 本客户端不写、只读的键（避免破坏共享设置）
READ_ONLY_KEYS = ("mail.passwordEnc", "mail.template.", "mail.ccAdmin.")


def current_values(settings):
    """当前邮件设置 → 界面值（布尔给 bool，其余给字符串）。"""
    values = {}
    for field in FIELDS:
        value = getattr(settings, field["attr"], None)
        if field["kind"] == "bool":
            values[field["key"]] = bool(value)
        elif field["kind"] == "int":
            values[field["key"]] = str(value if value is not None else 0)
        else:
            values[field["key"]] = "" if value is None else str(value)
    return values


def validate(values):
    """校验界面值；返回 ``(cleaned, errors, warnings)``。

    ``cleaned`` 是可直接写库的 ``{键: Python 值}``；有 ``errors`` 时不要写库。
    """
    errors = []
    warnings = []
    cleaned = {}

    for field in FIELDS:
        key = field["key"]
        label = field["label"]
        raw = values.get(key)
        kind = field["kind"]

        if kind == "bool":
            cleaned[key] = bool(raw)
            continue

        text = "" if raw is None else str(raw).strip()

        if kind == "int":
            if text == "":
                errors.append("%s：不能为空" % label)
                continue
            try:
                number = int(text)
            except ValueError:
                number = None
            if number is None:
                errors.append("%s：请填整数" % label)
                continue
            if number < field["min"] or number > field["max"]:
                errors.append("%s：必须在 %d–%d 之间" % (label, field["min"], field["max"]))
                continue
            cleaned[key] = number
            continue

        if kind == "choice":
            if text not in SECURITY_CODES:
                errors.append("%s：只能是 %s" % (label, " / ".join(SECURITY_CODES)))
                continue
            cleaned[key] = text
            continue

        # text：与主程序一致，空白视作「没有这一项」（写 NULL）
        cleaned[key] = text or None

    # 启用邮件时，服务器与发件地址必须齐（主程序的 IsReady 也是这两条）
    if cleaned.get("mail.enabled"):
        if not cleaned.get("mail.host"):
            errors.append("启用邮件功能时必须填 SMTP 服务器")
        if not mail_service.is_valid_address(cleaned.get("mail.fromAddress")):
            errors.append("启用邮件功能时必须填有效的发件地址")
    elif cleaned.get("mail.fromAddress") and not mail_service.is_valid_address(cleaned.get("mail.fromAddress")):
        errors.append("发件地址格式不正确（形如 user@example.com）")

    if cleaned.get("mail.ccList"):
        invalid = [
            item
            for item in mail_service.split_list(cleaned["mail.ccList"])
            if not mail_service.is_valid_address(item)
        ]
        if invalid:
            warnings.append("固定抄送里这些不是邮箱，会被忽略：%s" % "、".join(invalid))

    if cleaned.get("mail.useDefaultCredentials"):
        warnings.append("「Windows 集成验证」本客户端不支持（标准库没有 NTLM），发信会被跳过；请改用账号密码。")

    # 去重是两端共用的防重复发信机制（mail_logs + DedupeDays）：关掉它，两端都跑提醒就会重复发信
    if cleaned.get("mail.warningEnabled") and compat.to_int(cleaned.get("mail.dedupeDays"), 0) <= 0:
        warnings.append(
            "去重天数为 0：共享的 mail_logs 去重被关闭。如果主程序与 XP 端都会发到期提醒，"
            "同一封提醒会发两次；建议至少填 1。"
        )

    return cleaned, errors, warnings


def describe(cleaned):
    """一行行说明将要写入的内容（不含口令）。"""
    lines = []
    for field in FIELDS:
        key = field["key"]
        value = cleaned.get(key)
        if field["kind"] == "bool":
            text = "是" if value else "否"
        elif value is None:
            text = "(空)"
        else:
            text = str(value)
        lines.append("%s = %s" % (field["label"], text))
    return lines


def save(store, values):
    """校验并写入 ``app_settings``；返回 ``(cleaned, errors, warnings)``。

    有 ``errors`` 时**不写任何键**（多键一次事务，避免写一半）。
    """
    cleaned, errors, warnings = validate(values)
    if errors:
        return cleaned, errors, warnings
    store.set_many(cleaned)
    return cleaned, errors, warnings


def read_only_note():
    """界面上提示「哪些键不由本客户端维护」。"""
    return "口令（已加密）与邮件模板、抄送管理员开关由主程序维护，本客户端只读。"


def summary_line(settings):
    """设置摘要（界面上显示当前生效配置）。"""
    return "%s；安全方式 %s；提前 %d 天；含逾期 %s；去重 %d 天；口令来源 %s" % (
        ("%s:%d" % (settings.host, settings.port)) if settings.host else "(未配置服务器)",
        settings.security,
        settings.warning_days_before,
        "是" if settings.warning_include_overdue else "否",
        settings.dedupe_days,
        settings.password_source,
    )


def compare(old_settings, cleaned):
    """返回与旧设置相比真正变了的键（用于保存后提示改了什么）。"""
    changed = []
    for field in FIELDS:
        key = field["key"]
        new_value = cleaned.get(key)
        old_value = getattr(old_settings, field["attr"], None)
        if field["kind"] == "bool":
            old_normalized = bool(old_value)
            new_normalized = bool(new_value)
        elif field["kind"] == "int":
            old_normalized = compat.to_int(old_value, 0)
            new_normalized = compat.to_int(new_value, 0)
        else:
            old_normalized = None if old_value is None or str(old_value).strip() == "" else str(old_value)
            new_normalized = None if new_value is None or str(new_value).strip() == "" else str(new_value)
        if old_normalized != new_normalized:
            changed.append(key)
    return changed
