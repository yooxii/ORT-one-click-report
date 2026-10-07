# -*- coding: utf-8 -*-
"""tkinter 界面：登录 → 主窗口 → 领用与计划 / 邮件设置 / 环境自检。

界面刻意做得很朴素（需求方明确“UI 随意”），把力气留给“与主程序读写一致”这件事。
本里程碑的界面只做到：

- 登录（含记住登录）
- 领用表 / 计划表**只读列表**：搜索、刷新；
- 邮件设置查看、本机 SMTP 口令、到期提醒演练；
- 环境自检窗口。

编辑类操作（领退单新增/编辑、计划增改删）在里程碑 3 实现，邮件设置写库在里程碑 4 实现。
"""

import datetime
import tkinter as tk
from tkinter import messagebox, ttk

from .. import fatal
from ..db import repositories
from ..services import mail as mail_service
from ..services import mail_config
from ..services import plan_rules
from ..services import plans as plans_service
from ..version import APP_NAME, VERSION, BUILD_STAGE, TARGET_OS, TARGET_PYTHON

REQUISITION_COLUMNS = (
    ("RequisitionDate", "领用日期", 130),
    ("RequisitionNo", "领用单号", 120),
    ("ModelName", "机种", 140),
    ("WorkOrder", "工令", 130),
    ("OutQty", "领用数", 60),
    ("Disposition", "单体去向", 80),
    ("ReturnRtOrder", "回线RT工令", 120),
    ("Remark", "备注", 200),
)

PLAN_COLUMNS = (
    ("JobNo", "工作編號", 120),
    ("ModelName", "机种", 150),
    ("TestItem", "测试项目", 140),
    ("Stage", "阶段", 70),
    ("Owner", "负责人", 80),
    ("StartDate", "开始日期", 130),
    ("EndDate", "结束日期", 130),
    ("Status", "完成状况", 90),
)


class Application(object):
    def __init__(self, context):
        self.context = context
        self.root = tk.Tk()
        self.root.title("%s v%s" % (APP_NAME, VERSION))
        self.root.geometry("820x520")
        self.root.minsize(720, 420)
        # tkinter 回调里的异常默认只打到 stderr（窗口子系统看不到）→ 交给统一兜底
        self.root.report_callback_exception = self._on_callback_error
        self._apply_default_font()
        self.root.withdraw()

    def _on_callback_error(self, exc_type, exc_value, tb):
        fatal.handle(exc_type, exc_value, tb, self.context.app_directory)

    # ---------------------------------------------------------------- 启动

    def run(self):
        if not LoginDialog(self.root, self.context).show():
            self.root.destroy()
            return 0
        self._build_main()
        self.root.deiconify()
        self.root.mainloop()
        return 0

    def _apply_default_font(self):
        """XP 上默认字体偏小；统一放大一点，中文才不会挤在一起。"""
        try:
            default_font = ("Microsoft YaHei UI", 10)
            self.root.option_add("*Font", default_font)
        except Exception:
            pass

    def _build_main(self):
        frame = ttk.Frame(self.root, padding=12)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text=APP_NAME, font=("Microsoft YaHei UI", 16, "bold")).pack(anchor="w")
        ttk.Label(frame, text="v%s · %s" % (VERSION, BUILD_STAGE)).pack(anchor="w", pady=(2, 10))

        buttons = ttk.Frame(frame)
        buttons.pack(anchor="w", pady=(4, 12))
        ttk.Button(buttons, text="领用和计划", width=16, command=self.open_plans).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="邮件设置", width=16, command=self.open_mail).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="环境自检", width=16, command=self.open_selftest).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="注销", width=10, command=self.logout).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="退出", width=10, command=self.root.destroy).pack(side="left")

        info = ttk.LabelFrame(frame, text="运行状态", padding=10)
        info.pack(fill="both", expand=True)
        self._status_text = tk.Text(info, height=12, wrap="word")
        scroll = ttk.Scrollbar(info, orient="vertical", command=self._status_text.yview)
        self._status_text.configure(yscrollcommand=scroll.set)
        self._status_text.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self._refresh_status()

    def _refresh_status(self):
        context = self.context
        user = context.auth.current_user
        lines = [
            "当前用户：%s（%s）" % (context.auth.display_name(), ",".join(context.auth.roles) or "无角色"),
            "数据目录：%s（来源：%s）" % (context.data_dir, context.data_source),
            "数据库：%s" % context.database.describe(),
            "邮件：%s" % context.mail_settings.describe(),
            "目标环境：%s + Python %s" % (TARGET_OS, TARGET_PYTHON),
            "",
            "提示：本客户端支持领用表与计划表的新增/编辑/删除（改动都会写变更日志）；",
            "报告生成、审核、管理端、流程与索引不在本客户端范围内。",
            "带 * 为必填；日期格式 yyyy-MM-dd；双击表格行可直接编辑。",
        ]
        if user is not None and not user.email:
            lines.append("注意：当前账号没有邮箱，将收不到到期提醒。")
        self._status_text.configure(state="normal")
        self._status_text.delete("1.0", "end")
        self._status_text.insert("1.0", "\n".join(lines))
        self._status_text.configure(state="disabled")

    # ---------------------------------------------------------------- 各窗口

    def open_plans(self):
        PlansWindow(self.root, self.context)

    def open_mail(self):
        MailWindow(self.root, self.context, on_saved=self._refresh_status)

    def open_selftest(self):
        SelftestWindow(self.root, self.context)

    def logout(self):
        self.context.auth.logout()
        self.root.destroy()
        # 重新走一遍登录流程
        app = Application(self.context)
        app.run()


class LoginDialog(object):
    def __init__(self, parent, context):
        self.context = context
        self.window = tk.Toplevel(parent)
        self.window.title("登录 · %s" % APP_NAME)
        self.window.resizable(False, False)
        # 这里刻意**不**调用 transient(parent)：登录时主窗口（parent）是 withdraw 状态，
        # 而 Tk 在 Windows 上会把 master 已隐藏的 transient 窗口一起保持隐藏——
        # 实测（Tk 8.6，3.4 与 3.13 两个解释器）transient + withdrawn master 的窗口
        # winfo_viewable() 恒为 0，表现就是「进程活着、屏幕上什么都没有」。
        self.ok = False
        self._build()
        self.window.protocol("WM_DELETE_WINDOW", self._cancel)
        self.window.grab_set()

    def _build(self):
        frame = ttk.Frame(self.window, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=APP_NAME, font=("Microsoft YaHei UI", 14, "bold")).grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(frame, text="v%s · %s" % (VERSION, TARGET_OS)).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 12))

        ttk.Label(frame, text="用户名").grid(row=2, column=0, sticky="e", padx=(0, 8), pady=4)
        self.username = tk.StringVar()
        username_entry = ttk.Entry(frame, textvariable=self.username, width=28)
        username_entry.grid(row=2, column=1, sticky="we", pady=4)

        ttk.Label(frame, text="密码").grid(row=3, column=0, sticky="e", padx=(0, 8), pady=4)
        self.password = tk.StringVar()
        password_entry = ttk.Entry(frame, textvariable=self.password, show="*", width=28)
        password_entry.grid(row=3, column=1, sticky="we", pady=4)

        self.remember = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame, text="记住登录（本机 DPAPI 加密）", variable=self.remember).grid(
            row=4, column=1, sticky="w", pady=(4, 10)
        )

        buttons = ttk.Frame(frame)
        buttons.grid(row=5, column=0, columnspan=2, sticky="e")
        ttk.Button(buttons, text="登录", width=12, command=self._login).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="取消", width=10, command=self._cancel).pack(side="left")

        remembered = self.context.auth.load_remembered()
        if remembered:
            self.username.set(remembered[0])
            self.password.set(remembered[1])
            self.remember.set(True)
            password_entry.focus_set()
        else:
            username_entry.focus_set()

        self.window.bind("<Return>", lambda event: self._login())
        self.window.bind("<Escape>", lambda event: self._cancel())

    def _login(self):
        result = self.context.auth.login(self.username.get().strip(), self.password.get())
        if not result.ok:
            messagebox.showerror("登录失败", result.message, parent=self.window)
            return
        if self.remember.get():
            if not self.context.auth.remember():
                messagebox.showwarning("记住登录失败", "无法保存登录凭据（DPAPI 不可用）。", parent=self.window)
        else:
            self.context.auth.clear_remembered()
        if result.passwordless:
            messagebox.showinfo(
                "提示",
                "该账号还没有设置密码，本次按用户名直接登录。\n请在主程序里为用户设置密码。",
                parent=self.window,
            )
        self.ok = True
        self.window.destroy()

    def _cancel(self):
        self.ok = False
        self.window.destroy()

    def show(self):
        self.window.wait_window()
        return self.ok


class FormDialog(object):
    """通用表单对话框：tkinter 没有现成的表单控件，这里按字段定义生成。

    字段定义 ``(键, 标签, 类型, 选项)``，类型取 ``text`` / ``multiline`` / ``date`` / ``combo``。
    ``date`` 用文本框收 ``yyyy-MM-dd``（标准库没有日期选择器），留空表示不填。
    """

    def __init__(self, parent, title, fields, initial=None, hint=None, width=46):
        self.fields = fields
        self.initial = initial or {}
        self.values = None
        self._inputs = {}
        self.window = tk.Toplevel(parent)
        self.window.title(title)
        self.window.transient(parent)
        self.window.resizable(False, False)
        self._build(hint, width)
        self.window.protocol("WM_DELETE_WINDOW", self._cancel)
        self.window.bind("<Escape>", lambda event: self._cancel())
        self.window.grab_set()

    def _build(self, hint, width):
        frame = ttk.Frame(self.window, padding=12)
        frame.pack(fill="both", expand=True)
        row = 0
        if hint:
            ttk.Label(frame, text=hint, foreground="#555555", wraplength=width * 8).grid(
                row=row, column=0, columnspan=2, sticky="w", pady=(0, 8)
            )
            row += 1
        for key, label, kind, options in self.fields:
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="e", padx=(0, 8), pady=3)
            value = self.initial.get(key, "")
            text_value = "" if value is None else str(value)
            if kind == "multiline":
                widget = tk.Text(frame, width=width, height=3, wrap="word")
                widget.insert("1.0", text_value)
            elif kind == "combo":
                widget = ttk.Combobox(frame, values=list(options or ()), width=width - 3)
                widget.set(text_value)
            else:
                variable = tk.StringVar(value=text_value)
                widget = ttk.Entry(frame, textvariable=variable, width=width)
                self._inputs[key + "::var"] = variable
            widget.grid(row=row, column=1, sticky="we", pady=3)
            self._inputs[key] = widget
            row += 1

        buttons = ttk.Frame(frame)
        buttons.grid(row=row, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(buttons, text="保存", width=10, command=self._confirm).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="取消", width=8, command=self._cancel).pack(side="left")

    def read(self):
        result = {}
        for key, _label, kind, _options in self.fields:
            widget = self._inputs[key]
            if kind == "multiline":
                text = widget.get("1.0", "end").strip()
            elif kind == "combo":
                text = widget.get().strip()
            else:
                text = self._inputs[key + "::var"].get().strip()
            result[key] = text
        return result

    def _confirm(self):
        self.values = self.read()
        self.window.destroy()

    def _cancel(self):
        self.values = None
        self.window.destroy()

    def show(self):
        self.window.wait_window()
        return self.values


class TableTab(ttk.Frame):
    """一个表格页签（领用表 / 计划表共用）：搜索、增删改，改动都写变更日志。"""

    def __init__(self, parent, context, kind):
        ttk.Frame.__init__(self, parent, padding=8)
        self.context = context
        self.kind = kind  # requisitions / plans
        self.columns = REQUISITION_COLUMNS if kind == "requisitions" else PLAN_COLUMNS
        self.column_keys = tuple(item[0] for item in self.columns)
        self.sort_column = None
        self.sort_desc = False
        self._build()
        self.reload()

    def _build(self):
        toolbar = ttk.Frame(self)
        toolbar.pack(fill="x", pady=(0, 6))
        ttk.Label(toolbar, text="搜索").pack(side="left")
        self.keyword = tk.StringVar()
        entry = ttk.Entry(toolbar, textvariable=self.keyword, width=30)
        entry.pack(side="left", padx=6)
        entry.bind("<Return>", lambda event: self.reload())
        ttk.Button(toolbar, text="查询", command=self.reload).pack(side="left")
        ttk.Button(toolbar, text="刷新", command=self.reload).pack(side="left", padx=6)
        ttk.Button(toolbar, text="新增", command=self._add).pack(side="left", padx=(12, 0))
        ttk.Button(toolbar, text="编辑", command=self._edit).pack(side="left", padx=6)
        ttk.Button(toolbar, text="删除", command=self._delete).pack(side="left")
        self.count_label = ttk.Label(toolbar, text="")
        self.count_label.pack(side="left", padx=12)

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)
        keys = [item[0] for item in self.columns]
        self.tree = ttk.Treeview(container, columns=keys, show="headings", height=16)
        for key, title, width in self.columns:
            # 点列头排序：升序 → 降序 → 取消（回默认 Id 倒序）
            self.tree.heading(key, text=title, command=lambda column=key: self._sort_by(column))
            self.tree.column(key, width=width, anchor="w", stretch=True)
        scroll_y = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
        scroll_x = ttk.Scrollbar(container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="we")
        container.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)
        self.tree.bind("<Double-1>", self._on_double_click)

    def _on_double_click(self, event):
        self._edit()

    # ---------------------------------------------------------------- 排序

    def _sort_by(self, column):
        """点列头：第一次升序，再点降序，第三次回到默认（Id 倒序）。"""
        if self.sort_column != column:
            self.sort_column = column
            self.sort_desc = False
        elif not self.sort_desc:
            self.sort_desc = True
        else:
            self.sort_column = None
            self.sort_desc = False
        self._update_headings()
        self.reload()

    def _update_headings(self):
        for key, title, _width in self.columns:
            marker = ""
            if key == self.sort_column:
                marker = " ▼" if self.sort_desc else " ▲"
            self.tree.heading(key, text=title + marker)

    def _order_clause(self):
        """当前排序对应的 ORDER BY（列名走白名单，界面字符串不会拼进 SQL）。"""
        if not self.sort_column:
            return "Id DESC"
        return repositories.order_by_clause(self.column_keys, self.sort_column, self.sort_desc)

    # ---------------------------------------------------------------- 增删改

    def _operator(self):
        user = self.context.auth.current_user if self.context.auth is not None else None
        return user.username if user is not None else ""

    def _service(self):
        service = (
            self.context.requisition_service if self.kind == "requisitions" else self.context.plan_service
        )
        return service.with_operator(self._operator())

    def _repository(self):
        return (
            self.context.repositories.requisitions
            if self.kind == "requisitions"
            else self.context.repositories.plans
        )

    def _is_requisition(self):
        return self.kind == "requisitions"

    def _title(self):
        return "领退单" if self._is_requisition() else "计划"

    def _fields(self):
        """字段清单来自服务层（那边可脱离界面测试）。"""
        if self._is_requisition():
            return plans_service.requisition_form_fields()
        return plans_service.plan_form_fields(self.context.repositories.lookups)

    def _date_fields(self):
        if self._is_requisition():
            return plans_service.REQUISITION_DATE_FIELDS
        return plans_service.PLAN_DATE_FIELDS

    def _defaults(self):
        today = datetime.date.today().strftime("%Y-%m-%d")
        if self._is_requisition():
            return {"RequisitionDate": today, "Disposition": plan_rules.DISPOSITION_STOCK_IN}
        return {"StartDate": today, "Status": "Ongoing"}

    def _to_values(self, raw):
        """界面字符串 → 服务需要的值（日期转 datetime，空文本转 None）。"""
        keep_empty = ("Remark",) if self._is_requisition() else ()
        return plans_service.form_values(raw, self._date_fields(), keep_empty)

    def _from_row(self, row):
        return plans_service.form_initial(self._fields(), row, self._date_fields())

    def _selected_id(self):
        selection = self.tree.selection()
        if not selection:
            return None
        tags = self.tree.item(selection[0], "tags")
        if not tags:
            return None
        try:
            return int(tags[0])
        except (TypeError, ValueError):
            return None

    def _add(self):
        hint = "带 * 为必填。日期格式 yyyy-MM-dd；回线RT工令形如 RTAH260901。"
        raw = FormDialog(self.winfo_toplevel(), "新增" + self._title(), self._fields(), self._defaults(), hint).show()
        if raw is None:
            return
        self._apply(self._service().save(self._to_values(raw)))

    def _edit(self, record_id=None):
        if record_id is None:
            record_id = self._selected_id()
        if record_id is None:
            messagebox.showinfo("提示", "请先在表格里选择一行。", parent=self.winfo_toplevel())
            return
        row = self._repository().get(record_id)
        if row is None:
            messagebox.showwarning("提示", "该记录已不存在，列表将刷新。", parent=self.winfo_toplevel())
            self.reload()
            return
        raw = FormDialog(
            self.winfo_toplevel(), "编辑" + self._title(), self._fields(), self._from_row(row)
        ).show()
        if raw is None:
            return
        self._apply(self._service().save(self._to_values(raw), record_id))

    def _delete(self):
        record_id = self._selected_id()
        if record_id is None:
            messagebox.showinfo("提示", "请先在表格里选择一行。", parent=self.winfo_toplevel())
            return
        if not messagebox.askyesno(
            "删除确认", "确定删除选中的%s吗？\n删除会记入变更日志。" % self._title(), parent=self.winfo_toplevel()
        ):
            return
        self._apply(self._service().delete(record_id))

    def _apply(self, result):
        if not result.ok:
            title = "校验未通过" if result.errors else "操作失败"
            messagebox.showerror(title, result.error_text() or result.message, parent=self.winfo_toplevel())
            return
        messagebox.showinfo("完成", result.message or "已保存", parent=self.winfo_toplevel())
        self.reload()

    def reload(self):
        keyword = (self.keyword.get() or "").strip()
        keys = [item[0] for item in self.columns]
        try:
            # 查询统一走仓储（数据层只有一处 SQL），界面只负责展示
            repository = (
                self.context.repositories.requisitions if self.kind == "requisitions" else self.context.repositories.plans
            )
            rows = repository.list(keyword=keyword or None, limit=500, order=self._order_clause())
        except Exception as exc:
            messagebox.showerror("读取失败", str(exc), parent=self.winfo_toplevel())
            return

        self.tree.delete(*self.tree.get_children())
        for row in rows:
            values = []
            for key in keys:
                value = row[key]
                values.append("" if value is None else str(value))
            # Id 存进 tags，编辑/删除时用来定位记录
            self.tree.insert("", "end", values=tuple(values), tags=(str(row["Id"]),))
        self.count_label.configure(text="共 %d 行（最多显示 500 行）" % len(rows))


class PlansWindow(object):
    def __init__(self, parent, context):
        self.window = tk.Toplevel(parent)
        self.window.title("领用和计划")
        self.window.geometry("1080x620")
        notebook = ttk.Notebook(self.window)
        notebook.pack(fill="both", expand=True)
        notebook.add(TableTab(notebook, context, "requisitions"), text="领用表")
        notebook.add(TableTab(notebook, context, "plans"), text="计划表")


class MailWindow(object):
    """邮件设置：读写 ``app_settings`` 的 ``mail.*`` 键 + 本机 DPAPI 口令 + 测试发送。

    字段、取值范围与写出的文本格式与主程序设置窗口一致；口令（``mail.passwordEnc``）与
    邮件模板、抄送管理员开关**不在本窗口维护**（见 ``services/mail_config.py``）。
    """

    def __init__(self, parent, context, on_saved=None):
        self.context = context
        self.on_saved = on_saved
        self.vars = {}
        self._security_labels = dict((code, label) for code, label in mail_config.SECURITY_OPTIONS)
        self._security_codes = dict((label, code) for code, label in mail_config.SECURITY_OPTIONS)
        self.window = tk.Toplevel(parent)
        self.window.title("邮件设置")
        # 不写死 geometry：让 Tk 按内容算尺寸，避免字段多的时候按钮/输出区被裁掉
        self._build()
        self._load()

    # ---------------------------------------------------------------- 装配

    def _build(self):
        frame = ttk.Frame(self.window, padding=10)
        frame.pack(fill="both", expand=True)

        self.status = ttk.Label(frame, text="", wraplength=720)
        self.status.pack(anchor="w", pady=(0, 6))

        columns = ttk.Frame(frame)
        columns.pack(fill="x")

        smtp = ttk.LabelFrame(columns, text="SMTP（写入共享设置，主程序同样读取）", padding=8)
        smtp.pack(side="left", fill="both", expand=True)
        switches = ttk.LabelFrame(columns, text="开关与提醒", padding=8)
        switches.pack(side="left", fill="both", expand=True, padx=(8, 0))

        row = 0
        for field in mail_config.FIELDS:
            if field["kind"] == "bool":
                continue
            row = self._add_field(smtp, field, row)
        row = 0
        for field in mail_config.FIELDS:
            if field["kind"] != "bool":
                continue
            row = self._add_field(switches, field, row)

        local = ttk.LabelFrame(frame, text="本机 SMTP 口令（DPAPI 加密，只存本机）", padding=8)
        local.pack(fill="x", pady=8)
        ttk.Label(local, text=mail_config.read_only_note(), wraplength=760).pack(anchor="w")
        password_row = ttk.Frame(local)
        password_row.pack(fill="x", pady=(4, 0))
        self.password = tk.StringVar()
        ttk.Entry(password_row, textvariable=self.password, show="*", width=30).pack(side="left")
        ttk.Button(password_row, text="保存到本机", command=self._save_password).pack(side="left", padx=8)
        self.password_state = ttk.Label(password_row, text="")
        self.password_state.pack(side="left")

        test_row = ttk.Frame(frame)
        test_row.pack(fill="x")
        ttk.Label(test_row, text="测试收件地址").pack(side="left")
        self.test_to = tk.StringVar()
        ttk.Entry(test_row, textvariable=self.test_to, width=28).pack(side="left", padx=6)
        ttk.Label(test_row, text="（留空用发件地址）").pack(side="left")

        actions = ttk.Frame(frame)
        actions.pack(fill="x", pady=6)
        ttk.Button(actions, text="保存并生效", width=12, command=self._save).pack(side="left")
        ttk.Button(actions, text="测试发送（演练）", command=lambda: self._test_send(False)).pack(side="left", padx=6)
        ttk.Button(actions, text="测试发送（真实）", command=lambda: self._test_send(True)).pack(side="left")
        ttk.Button(actions, text="到期提醒演练", command=self._dry_run_reminder).pack(side="left", padx=6)
        ttk.Button(actions, text="刷新发送记录", command=self._refresh_logs).pack(side="left")

        self.output = tk.Text(frame, height=6, wrap="word")
        scroll = ttk.Scrollbar(frame, orient="vertical", command=self.output.yview)
        self.output.configure(yscrollcommand=scroll.set)
        self.output.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

    def _add_field(self, parent, field, row):
        key = field["key"]
        kind = field["kind"]
        if kind == "bool":
            variable = tk.BooleanVar(value=False)
            ttk.Checkbutton(parent, text=field["label"], variable=variable).grid(
                row=row, column=0, columnspan=2, sticky="w", pady=1
            )
        else:
            ttk.Label(parent, text=field["label"]).grid(row=row, column=0, sticky="e", padx=(0, 8), pady=2)
            variable = tk.StringVar()
            if kind == "choice":
                widget = ttk.Combobox(
                    parent,
                    textvariable=variable,
                    state="readonly",
                    width=12,
                    values=[label for _code, label in field["options"]],
                )
            else:
                widget = ttk.Entry(parent, textvariable=variable, width=26)
            widget.grid(row=row, column=1, sticky="we", pady=2)
        self.vars[key] = variable
        return row + 1

    # ---------------------------------------------------------------- 读 / 写

    def _load(self):
        """把当前生效的设置填进控件（保存后与打开时都用它）。"""
        values = mail_config.current_values(self.context.mail_settings)
        for field in mail_config.FIELDS:
            key = field["key"]
            value = values.get(key)
            if field["kind"] == "bool":
                self.vars[key].set(bool(value))
            elif field["kind"] == "choice":
                self.vars[key].set(self._security_labels.get(value, self._security_labels["None"]))
            else:
                self.vars[key].set("" if value is None else str(value))
        settings = self.context.mail_settings
        ready, reason = self.context.mail.is_ready()
        self.status.configure(
            text="当前生效：%s\n发信条件：%s"
            % (mail_config.summary_line(settings), "可发信" if ready else "未就绪 —— %s" % reason)
        )
        self.password_state.configure(text="当前口令来源：%s" % settings.password_source)

    def _collect(self):
        values = {}
        for field in mail_config.FIELDS:
            key = field["key"]
            raw = self.vars[key].get()
            if field["kind"] == "bool":
                values[key] = bool(raw)
            elif field["kind"] == "choice":
                values[key] = self._security_codes.get(raw, "None")
            else:
                values[key] = raw
        return values

    def _save(self):
        before = self.context.mail_settings
        cleaned, errors, warnings = mail_config.save(self.context.settings_store, self._collect())
        if errors:
            for line in errors:
                self._write("校验未通过：" + line)
            messagebox.showerror("校验未通过", "\n".join(errors), parent=self.window)
            return
        changed = mail_config.compare(before, cleaned)
        self.context.reload_mail()
        self._load()
        for line in warnings:
            self._write("提醒：" + line)
        self._write(
            "已保存到 app_settings：%s" % ("、".join(changed) if changed else "（内容与原来一致）")
        )
        if self.on_saved:
            self.on_saved()

    # ---------------------------------------------------------------- 口令与发送

    def _write(self, text):
        self.output.configure(state="normal")
        self.output.insert("end", text + "\n")
        self.output.see("end")
        self.output.configure(state="disabled")

    def _save_password(self):
        password = self.password.get()
        if self.context.credentials.save_mail_password(password):
            self._write("本机口令已保存。" if password else "本机口令已清除。")
            self.password.set("")
            self.context.reload_mail()
            self._load()
            if self.on_saved:
                self.on_saved()
        else:
            messagebox.showerror("保存失败", "DPAPI 不可用，无法加密保存。", parent=self.window)

    def _test_send(self, real):
        recipient = (self.test_to.get() or "").strip() or self.context.mail_settings.from_address
        if not mail_service.is_valid_address(recipient):
            self._write("请先填一个有效的测试收件地址（或在上面配置发件地址）。")
            return
        if real and not messagebox.askyesno(
            "确认发送", "确定现在真的发一封测试邮件到 %s 吗？" % recipient, parent=self.window
        ):
            return
        result = self.context.mail.send_test(recipient, dry_run=not real)
        self._write(
            "测试发送（%s）→ %s：%s" % ("真实发送" if real else "演练，不发信", ";".join(result.recipients) or recipient, result.message)
        )
        self._refresh_logs()

    def _dry_run_reminder(self):
        from ..context import run_reminder

        self._write(run_reminder(self.context, dry_run=True))

    def _refresh_logs(self):
        try:
            rows = self.context.mail.recent_logs(20)
        except Exception as exc:
            self._write("读取发送记录失败：%s" % exc)
            return
        if not rows:
            self._write("（还没有发送记录；也可能是数据文件夹里没有 mail_logs 表）")
            return
        self._write("最近 %d 条发送记录：" % len(rows))
        for row in rows:
            self._write(
                "  %s [%s] %s → %s：%s%s"
                % (
                    row["CreatedAt"],
                    row["Kind"],
                    row["RefKey"] or "-",
                    row["Recipients"] or "-",
                    "成功" if row["Success"] else "失败",
                    "" if row["Success"] else "（%s）" % (row["Error"] or ""),
                )
            )


class SelftestWindow(object):
    def __init__(self, parent, context):
        self.window = tk.Toplevel(parent)
        self.window.title("环境自检")
        self.window.geometry("760x520")
        frame = ttk.Frame(self.window, padding=12)
        frame.pack(fill="both", expand=True)
        text = tk.Text(frame, wrap="word")
        scroll = ttk.Scrollbar(frame, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=scroll.set)
        text.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        text.insert("1.0", context.selftest_text())
        text.configure(state="disabled")


def _smoke_text(steps):
    lines = ["界面装配自检（只造窗口，不进入主循环）", ""]
    for name, ok, detail in steps:
        lines.append("[%s] %s%s" % ("通过" if ok else "失败", name, "" if ok else "：%s" % detail))
    return "\n".join(lines)


def ui_smoke(context):
    """无人值守的界面装配自检：把主要窗口造出来、确认**看得见**再销毁，返回 ``(ok, 文本)``。

    窗口子系统下没法人工点击，这条命令用来在目标机上证明「界面确实能显示出来」，
    而不是靠双击碰运气。只检查「造得出来」是不够的：Tk 会把 master 已 withdraw 的
    transient 窗口一起隐藏，这种问题只有查 ``winfo_viewable()`` 才发现得了。
    """
    steps = []

    def _step(name, func):
        try:
            func()
        except Exception as exc:
            steps.append((name, False, "%s: %s" % (exc.__class__.__name__, exc)))
        else:
            steps.append((name, True, ""))

    def _pump(window):
        """让 Tk 把窗口真正映射出来（不跑 mainloop）。"""
        try:
            window.update_idletasks()
            window.update()
        except tk.TclError:
            pass

    def _require_viewable(window, name):
        _pump(window)
        if not window.winfo_viewable():
            raise AssertionError("%s没有显示出来（winfo_viewable() = 0）" % name)

    application = None
    try:
        try:
            application = Application(context)
        except Exception as exc:
            steps.append(("根窗口", False, "%s: %s" % (exc.__class__.__name__, exc)))
            return False, _smoke_text(steps)
        root = application.root

        def _login():
            dialog = LoginDialog(root, context)
            try:
                _require_viewable(dialog.window, "登录窗口")
                if not isinstance(dialog.username.get(), str) or not isinstance(dialog.password.get(), str):
                    raise AssertionError("登录窗口的输入框变量不可用")
            finally:
                dialog.window.destroy()

        def _main():
            application._build_main()
            root.deiconify()
            _require_viewable(root, "主窗口")

        def _tabs():
            for kind in ("requisitions", "plans"):
                tab = TableTab(root, context, kind)
                try:
                    tab.reload()
                finally:
                    tab.destroy()

        def _plans_window():
            window = PlansWindow(root, context)
            try:
                _require_viewable(window.window, "领用和计划窗口")
            finally:
                window.window.destroy()

        def _form():
            dialog = FormDialog(
                root, "界面自检", plans_service.requisition_form_fields(), {"RequisitionNo": "SMOKE-0001"}
            )
            try:
                _require_viewable(dialog.window, "表单对话框")
                values = dialog.read()
                if values.get("RequisitionNo") != "SMOKE-0001":
                    raise AssertionError("表单回读不一致：%r" % (values.get("RequisitionNo"),))
            finally:
                dialog.window.destroy()

        def _mail():
            window = MailWindow(root, context)
            try:
                _require_viewable(window.window, "邮件设置窗口")
                if not window.window.title():
                    raise AssertionError("邮件设置窗口没有标题")
            finally:
                window.window.destroy()

        def _selftest():
            window = SelftestWindow(root, context)
            try:
                _require_viewable(window.window, "环境自检窗口")
                if not window.window.title():
                    raise AssertionError("环境自检窗口没有标题")
            finally:
                window.window.destroy()

        _step("登录窗口", _login)
        _step("主窗口", _main)
        _step("领用表 / 计划表页签", _tabs)
        _step("领用和计划窗口", _plans_window)
        _step("表单对话框", _form)
        _step("邮件设置窗口", _mail)
        _step("环境自检窗口", _selftest)
    finally:
        if application is not None:
            try:
                application.root.destroy()
            except Exception:
                pass
    return all(item[1] for item in steps), _smoke_text(steps)


def run(context):
    """启动界面；返回进程退出码。"""
    try:
        return Application(context).run()
    except tk.TclError as exc:
        return fatal.report(
            "无法启动界面",
            "界面（tkinter/Tcl）启动失败：%s" % exc,
            context.app_directory,
            context.logger,
        )
