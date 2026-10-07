# -*- coding: utf-8 -*-
"""计划 / 领退的校验规则与自动编号。

逐条对齐主程序的：

- ``Utils/PlanValidation.cs``：工作编号、回线RT工令、状况、字典校验，以及状况归类；
- ``Services/PlanExcelService.cs`` 的 ``GenerateJobNo`` / ``GenerateReturnRtOrder``：
  序号取「当月已有编号末尾序号的最大值 + 1」（不是 COUNT+1，删过历史也不会重号）。

本模块是**纯函数**，不碰数据库：调用方把已有的编号清单传进来。
"""

import re

# --------------------------------------------------------------------------- 常量

#: 状况归类（对应 PlanStatusKind）
STATUS_ONGOING = "ongoing"
STATUS_PENDING = "pending"
STATUS_CLOSED = "closed"

#: 允许的状况值（对应 PlanValidation.ValidStatuses）
VALID_STATUSES = ("Ongoing", "Close", "Pending")

#: 单体去向（对应 RequisitionDispositionKind）
DISPOSITION_STOCK_IN = "入库"
DISPOSITION_SCRAP = "报废"
DISPOSITION_ALL = (DISPOSITION_STOCK_IN, DISPOSITION_SCRAP)

#: 工作编号：QRT/RT + 4 位年月 + 至少 2 位编号
JOB_NO_RE = re.compile(r"^(QRT|RT)(\d{4})(\d{2,})$", re.IGNORECASE)

#: 回线RT工令：RTAH + 4 位年月 + 至少 2 位编号
RETURN_RT_RE = re.compile(r"^RTAH(\d{4})(\d{2,})$", re.IGNORECASE)

MESSAGE_JOB_NO_FORMAT = "工作编号格式应为：QRT/RT + 4位年月 + 至少2位编号（如 RT260801）"
MESSAGE_JOB_NO_SEQ = "工作编号末尾编号必须从 01 开始"
MESSAGE_RETURN_RT_FORMAT = "回线RT工令格式应为：RTAH + 4位年月 + 至少2位编号（如 RTAH260901）"
MESSAGE_RETURN_RT_SEQ = "回线RT工令末尾编号必须从 01 开始"
MESSAGE_STATUS = "状况只能是：%s" % " / ".join(VALID_STATUSES)


# --------------------------------------------------------------------------- 校验


def _clean(value):
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def validate_job_no(value):
    """合法（或为空）返回 None，否则返回错误描述。"""
    text = _clean(value)
    if text is None:
        return None
    match = JOB_NO_RE.match(text)
    if not match:
        return MESSAGE_JOB_NO_FORMAT
    if int(match.group(3)) < 1:
        return MESSAGE_JOB_NO_SEQ
    return None


def validate_return_rt_order(value):
    """回线RT工令可以为空（报废无需回线），但不能是别的格式。"""
    text = _clean(value)
    if text is None:
        return None
    match = RETURN_RT_RE.match(text)
    if not match:
        return MESSAGE_RETURN_RT_FORMAT
    if int(match.group(2)) < 1:
        return MESSAGE_RETURN_RT_SEQ
    return None


def validate_status(value):
    text = _clean(value)
    if text is None:
        return None
    for allowed in VALID_STATUSES:
        if allowed.lower() == text.lower():
            return None
    return MESSAGE_STATUS


def validate_disposition(value):
    text = _clean(value)
    if text is None:
        return "请选择单体去向（%s / %s）" % DISPOSITION_ALL
    if text not in DISPOSITION_ALL:
        return "单体去向只能是：%s / %s" % DISPOSITION_ALL
    return None


def validate_in_catalog(value, catalog, column_name):
    """值必须在字典里（空值合法）；对应主程序 ``ValidateInCatalog``。"""
    text = _clean(value)
    if text is None or not catalog:
        return None
    for item in catalog:
        if item == text:
            return None
    return "%s [%s] 不在字典中，请先在主程序的“管理”模块中添加" % (column_name, text)


def status_kind(status):
    """状况归类（对应 ``PlanStatusKind.Of``）：认不出返回空字符串。"""
    text = _clean(status)
    if text is None:
        return ""
    lowered = text.lower()
    for token in ("ongoing", "进行中", "進行中", "測試中", "测试中"):
        if token in lowered:
            return STATUS_ONGOING
    for token in ("pending", "待测", "待測", "預排", "预排"):
        if token in lowered:
            return STATUS_PENDING
    for token in ("close", "结案", "結案", "已完成"):
        if token in lowered:
            return STATUS_CLOSED
    return ""


# --------------------------------------------------------------------------- 自动编号


def format_sequence(seq):
    """序号格式化：<100 补到两位，≥100 按实际位数展开。"""
    return str(seq) if seq >= 100 else "%02d" % seq


def max_trailing_sequence(values, pattern):
    """从一组编号里按正则取出末尾序号的最大值（不匹配的跳过）。"""
    regex = re.compile(pattern, re.IGNORECASE)
    maximum = 0
    for value in values or ():
        text = _clean(value)
        if text is None:
            continue
        match = regex.match(text)
        if not match:
            continue
        try:
            seq = int(match.group(1))
        except ValueError:
            continue
        if seq > maximum:
            maximum = seq
    return maximum


def generate_job_no(existing_job_nos, date, prefix="RT"):
    """生成工作编号 ``{prefix}{yyMM}{序号}``。

    RT 与 QRT **共享同一个月度序号**（与主程序一致）：两种前缀的记录一起参与取最大值。
    """
    ym = date.strftime("%y%m")
    use_prefix = (prefix or "RT").strip().upper() or "RT"
    pattern = r"^(?:QRT|RT)%s(\d+)$" % re.escape(ym)
    next_seq = max_trailing_sequence(existing_job_nos, pattern) + 1
    return "%s%s%s" % (use_prefix, ym, format_sequence(next_seq))


def generate_return_rt_order(existing_orders, date):
    """生成回线RT工令 ``RTAH{yyMM}{序号}``。"""
    ym = date.strftime("%y%m")
    pattern = r"^RTAH%s(\d+)$" % re.escape(ym)
    next_seq = max_trailing_sequence(existing_orders, pattern) + 1
    return "RTAH%s%s" % (ym, format_sequence(next_seq))
