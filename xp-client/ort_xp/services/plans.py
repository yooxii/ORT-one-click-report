# -*- coding: utf-8 -*-
"""计划与领退的编辑服务：校验 + 落库 + 变更日志。

校验顺序与错误措辞都对齐主程序的编辑窗口（``WindowRequisitionEdit`` /
``WindowPlanDirectEdit``）：先是必填，再是格式，最后是唯一性。
与主程序的差别只有一处：主程序是「弹第一个错误就返回」，这里把错误**收集成列表**
一次返回，界面上一次就能看全部问题；规则本身相同。

界面（``ui/app.py``）只负责收集输入与显示错误，业务判断都在这里，因此可以脱离界面测试。
"""

from .. import compat
from . import plan_rules

#: 领退单必填字段（顺序与主程序校验顺序一致）
REQUISITION_REQUIRED = (
    ("RequisitionDate", "请选择领用日期"),
    ("RequisitionNo", "请填写領料單据號"),
    ("ModelName", "请填写机种名"),
    ("OutQty", "请填写领用数量"),
    ("Rev", "请填写版本"),
    ("WorkOrder", "请填写工令"),
)

#: 计划表必填字段（备注在保存时也是必填，与主程序一致）
PLAN_REQUIRED = (
    ("TestItem", "请选择测试项目"),
    ("StartDate", "请填写开始时间"),
    ("Stage", "请选择阶段"),
    ("ModelName", "请填写机种名"),
    ("Remark", "请填写备注"),
)

#: 领退单里需要去掉首尾空白的文本列
REQUISITION_TEXT_FIELDS = (
    "RequisitionNo",
    "ModelName",
    "OutQty",
    "Rev",
    "WorkOrder",
    "DC",
    "LineNo",
    "SN",
    "SnFilePath",
    "Disposition",
    "ReturnRtOrder",
    "Remark",
)

#: 计划表里需要去掉首尾空白的文本列
PLAN_TEXT_FIELDS = (
    "JobNo",
    "ModelName",
    "TestItem",
    "Product",
    "Customer",
    "Stage",
    "SampleSize",
    "TestPeriod",
    "Owner",
    "Status",
    "UploadELab",
    "Remark",
)


def _present(value):
    if value is None:
        return False
    if isinstance(value, str):
        return value.strip() != ""
    return True


def _clean(value):
    if value is None:
        return None
    text = str(value).strip()
    return text or None


# --------------------------------------------------------------------------- 表单定义
#
# 表单字段放这里而不是界面里：界面只负责画控件，字段清单与「界面字符串 ↔ 服务值」的转换
# 都是纯数据逻辑，放在这儿就能脱离界面测试（tkinter 的窗口没法在无人值守下点）。

REQUISITION_DATE_FIELDS = ("RequisitionDate",)
PLAN_DATE_FIELDS = ("StartDate", "EndDate")

#: 精简版界面不提供的字段及原因（服务层仍可写，方便以后补界面）
FORM_EXEMPT_FIELDS = {
    "SnFilePath": "序列号附件文件（选择/上传文件）不在精简版范围内",
}


def requisition_form_fields():
    """领退单表单字段：``(键, 标签, 类型, 选项)``。"""
    return (
        ("RequisitionDate", "领用日期 *", "date", ()),
        ("RequisitionNo", "領料單據號 *", "text", ()),
        ("ModelName", "机种名 *", "text", ()),
        ("OutQty", "领用数量 *", "text", ()),
        ("Rev", "版本 *", "text", ()),
        ("WorkOrder", "工令 *", "text", ()),
        ("DC", "DC", "text", ()),
        ("LineNo", "线别", "text", ()),
        ("Disposition", "单体去向 *", "combo", plan_rules.DISPOSITION_ALL),
        ("ReturnRtOrder", "回线RT工令（入库必填）", "text", ()),
        ("SN", "序列号（每行一个）", "multiline", ()),
        ("Remark", "备注", "multiline", ()),
    )


def plan_form_fields(lookups):
    """计划表表单字段（测试项目/阶段/产品/客户用字典做下拉）。"""
    return (
        ("JobNo", "工作编号（留空自动生成）", "text", ()),
        ("ModelName", "机种名 *", "text", ()),
        ("TestItem", "测试项目 *", "combo", tuple(lookups.test_items())),
        ("Stage", "阶段 *", "combo", tuple(lookups.stages())),
        ("Product", "产品别", "combo", tuple(lookups.products())),
        ("Customer", "客户别", "combo", tuple(lookups.customers())),
        ("Owner", "负责人", "text", ()),
        ("SampleSize", "样品数", "text", ()),
        ("TestPeriod", "试验时间", "text", ()),
        ("StartDate", "开始日期 *", "date", ()),
        ("EndDate", "结束日期", "date", ()),
        ("Status", "完成状况", "combo", plan_rules.VALID_STATUSES),
        ("UploadELab", "上传 e-lab", "text", ()),
        ("Remark", "备注 *", "multiline", ()),
    )


def form_values(raw, date_fields, keep_empty=()):
    """界面字符串 → 服务需要的值：日期解析成 datetime，空文本转 None。

    ``keep_empty`` 里的字段保留空字符串（主程序的备注就是空串而不是 NULL）。
    """
    values = {}
    for key, value in raw.items():
        if key in date_fields:
            values[key] = compat.parse_datetime_text(value) if value else None
        elif key in keep_empty:
            values[key] = value if value is not None else ""
        else:
            values[key] = value if value != "" else None
    return values


def form_initial(fields, row, date_fields):
    """行数据 → 表单初值（日期显示成 ``yyyy-MM-dd``）。"""
    initial = {}
    for key, _label, _kind, _options in fields:
        value = row[key]
        if value is None:
            initial[key] = ""
        elif key in date_fields:
            parsed = compat.parse_datetime_text(value)
            initial[key] = parsed.strftime("%Y-%m-%d") if parsed is not None else str(value)
        else:
            initial[key] = str(value)
    return initial


class EditResult(object):
    """保存/删除的结果。"""

    def __init__(self, ok, errors=None, record_id=None, message=""):
        self.ok = ok
        self.errors = list(errors or [])
        self.record_id = record_id
        self.message = message

    def error_text(self):
        return "\n".join(self.errors)

    def __repr__(self):
        return "<EditResult %s %s>" % ("成功" if self.ok else "失败", self.error_text() or self.message)


class _EditService(object):
    def __init__(self, repositories, lookups=None, operator=""):
        self.repositories = repositories
        self.lookups = lookups if lookups is not None else repositories.lookups
        self.operator = operator or ""
        self.logger = None

    def with_operator(self, operator):
        """返回一个换了操作人的同名服务（界面按当前登录用户取）。"""
        return self.__class__(self.repositories, self.lookups, operator)

    # ---------------------------------------------------------------- 字典

    def _catalog(self, name):
        getter = getattr(self.lookups, name, None)
        if getter is None:
            return []
        try:
            return list(getter())
        except Exception:
            return []


class RequisitionService(_EditService):
    """领退单的校验与保存。"""

    def validate(self, values, record_id=None):
        errors = []
        for field, message in REQUISITION_REQUIRED:
            if not _present(values.get(field)):
                errors.append(message)

        disposition = _clean(values.get("Disposition"))
        disposition_error = plan_rules.validate_disposition(disposition)
        if disposition_error:
            errors.append(disposition_error)

        return_rt = _clean(values.get("ReturnRtOrder"))
        if disposition == plan_rules.DISPOSITION_STOCK_IN and return_rt is None:
            errors.append("单体去向为「入库」时必须填写回线RT工令")
        message = plan_rules.validate_return_rt_order(return_rt)
        if message:
            errors.append(message)
        elif return_rt and self.repositories.requisitions.exists_value("ReturnRtOrder", return_rt, record_id):
            errors.append("回线RT工令 [%s] 已存在" % return_rt)

        requisition_no = _clean(values.get("RequisitionNo"))
        if requisition_no and self.repositories.requisitions.requisition_no_exists(requisition_no, record_id):
            errors.append("領料單据號 [%s] 已存在" % requisition_no)
        return errors

    def payload(self, values):
        payload = {}
        for field in REQUISITION_TEXT_FIELDS:
            if field not in values:
                continue
            if field == "Remark":
                payload[field] = (values.get(field) or "").strip()
            else:
                payload[field] = _clean(values.get(field))
        if "RequisitionDate" in values:
            payload["RequisitionDate"] = values.get("RequisitionDate")
        return payload

    def save(self, values, record_id=None):
        errors = self.validate(values, record_id)
        if errors:
            return EditResult(False, errors)
        payload = self.payload(values)
        try:
            if record_id:
                changed = self.repositories.requisitions.update(record_id, payload, self.operator)
                message = "已保存修改" if changed else "内容没有变化，未写入"
                return EditResult(True, record_id=record_id, message=message)
            new_id = self.repositories.requisitions.insert(payload, self.operator)
            return EditResult(True, record_id=new_id, message="已新增领退记录")
        except Exception as exc:
            return EditResult(False, [str(exc)])

    def delete(self, record_id):
        try:
            if self.repositories.requisitions.delete(record_id, self.operator):
                return EditResult(True, record_id=record_id, message="已删除")
            return EditResult(False, ["记录不存在"])
        except Exception as exc:
            return EditResult(False, [str(exc)])

    # ---------------------------------------------------------------- 自动编号

    def existing_return_rt_orders(self, date):
        if not self.repositories.database.table_exists("requisitions"):
            return []
        ym = date.strftime("%y%m")
        rows = self.repositories.database.query(
            'SELECT "ReturnRtOrder" FROM "requisitions" WHERE "ReturnRtOrder" LIKE ?', ("RTAH" + ym + "%",)
        )
        return [row["ReturnRtOrder"] for row in rows]

    def next_return_rt_order(self, date):
        return plan_rules.generate_return_rt_order(self.existing_return_rt_orders(date), date)

    def next_job_no(self, date, prefix="RT"):
        """与主程序一致：领退侧建立计划时用 RT 前缀（RT/QRT 共享月度序号）。"""
        return plan_rules.generate_job_no(self._all_job_nos(), date, prefix)

    def _all_job_nos(self):
        if not self.repositories.database.table_exists("plans"):
            return []
        rows = self.repositories.database.query('SELECT "JobNo" FROM "plans" WHERE "JobNo" IS NOT NULL')
        return [row["JobNo"] for row in rows]


class PlanService(_EditService):
    """计划表的校验与保存。"""

    def validate(self, values, record_id=None):
        errors = []
        for field, message in PLAN_REQUIRED:
            if not _present(values.get(field)):
                errors.append(message)

        status_error = plan_rules.validate_status(values.get("Status"))
        if status_error:
            errors.append(status_error)

        job_no = _clean(values.get("JobNo"))
        if job_no is None and _present(values.get("StartDate")):
            job_no = self.next_job_no(values.get("StartDate"))
        job_no_error = plan_rules.validate_job_no(job_no)
        if job_no_error:
            errors.append(job_no_error)
        elif job_no and self.repositories.plans.job_no_exists(job_no, record_id):
            errors.append("工作编号 [%s] 已存在" % job_no)

        # 字典校验：字典为空时不做限制（允许在主程序尚未维护字典的库里使用）
        for field, catalog_name, column_name in (
            ("TestItem", "test_items", "测试项目"),
            ("Stage", "stages", "阶段"),
            ("Product", "products", "产品别"),
            ("Customer", "customers", "客户别"),
        ):
            catalog = self._catalog(catalog_name)
            message = plan_rules.validate_in_catalog(values.get(field), catalog, column_name)
            if message:
                errors.append(message)
        return errors

    def payload(self, values):
        payload = {}
        for field in PLAN_TEXT_FIELDS:
            if field in values:
                payload[field] = _clean(values.get(field))
        for field in ("StartDate", "EndDate", "UnitReturnDate"):
            if field in values:
                payload[field] = values.get(field)
        return payload

    def resolve_job_no(self, values):
        """填了就用手填的，留空则按开始日期生成 QRT{yyMM}{序号}（与主程序一致）。"""
        job_no = _clean(values.get("JobNo"))
        if job_no:
            return job_no
        start = values.get("StartDate")
        if not _present(start):
            return None
        return self.next_job_no(start)

    def next_job_no(self, date, prefix="QRT"):
        return plan_rules.generate_job_no(self._all_job_nos(), date, prefix)

    def _all_job_nos(self):
        if not self.repositories.database.table_exists("plans"):
            return []
        rows = self.repositories.database.query('SELECT "JobNo" FROM "plans" WHERE "JobNo" IS NOT NULL')
        return [row["JobNo"] for row in rows]

    def save(self, values, record_id=None):
        values = dict(values)
        if not _clean(values.get("JobNo")):
            values["JobNo"] = self.resolve_job_no(values)
        errors = self.validate(values, record_id)
        if errors:
            return EditResult(False, errors)
        payload = self.payload(values)
        try:
            if record_id:
                changed = self.repositories.plans.update(record_id, payload, self.operator)
                message = "已保存修改" if changed else "内容没有变化，未写入"
                return EditResult(True, record_id=record_id, message=message)
            new_id = self.repositories.plans.insert(payload, self.operator)
            return EditResult(True, record_id=new_id, message="已新增计划")
        except Exception as exc:
            return EditResult(False, [str(exc)])

    def delete(self, record_id):
        try:
            if self.repositories.plans.delete(record_id, self.operator):
                return EditResult(True, record_id=record_id, message="已删除")
            return EditResult(False, ["记录不存在"])
        except Exception as exc:
            return EditResult(False, [str(exc)])

    def suggest_product_customer(self, model_name):
        """按机种名给出产品别/客户别建议（对齐主程序：先查机种映射表，再按编码规则）。

        仅作界面上的自动填充，用户可以改。
        """
        model_name = _clean(model_name)
        if model_name is None:
            return None, None
        mapping = self.lookups.find_model_mapping(model_name)
        if mapping is not None:
            product = mapping["Product"] or None
            customer = mapping["Customer"] or None
            if product or customer:
                return product, customer
        return self.lookups.find_product_by_model(model_name), self.lookups.find_customer_by_model(model_name)
