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

import tkinter as tk
from tkinter import messagebox, ttk

from .. import compat
from ..version import APP_NAME, VERSION, BUILD_STAGE, TARGET_OS, TARGET_PYTHON

REQUISITION_COLUMNS = (
    ("RequisitionDate", "领用日期", 100),
    ("RequisitionNo", "领用单号", 110),
    ("ModelName", "机种", 140),
    ("WorkOrder", "工令", 110),
    ("OutQty", "领用数", 60),
    ("Disposition", "单体去向", 80),
    ("ReturnRtOrder", "回线RT工令", 110),
    ("Remark", "备注", 200),
)

PLAN_COLUMNS = (
    ("JobNo", "工作編號", 110),
    ("ModelName", "机种", 150),
    ("TestItem", "测试项目", 140),
    ("Stage", "阶段", 70),
    ("Owner", "负责人", 80),
    ("StartDate", "开始日期", 100),
    ("EndDate", "结束日期", 100),
    ("Status", "完成状况", 90),
)

EDIT_HINT = "编辑功能将在里程碑 3 实现（当前为只读列表）。"


class Application(object):
    def __init__(self, context):
        self.context = context
        self.root = tk.Tk()
        self.root.title("%s v%s" % (APP_NAME, VERSION))
        self.root.geometry("820x520")
        self.root.minsize(720, 420)
        self._apply_default_font()
        self.root.withdraw()

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
            "提示：本里程碑界面只读；编辑、审核、报告生成不在本客户端范围内。",
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
        self.window.transient(parent)
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
        ttk.Entry(frame, textvariable=self.username, width=28).grid(row=2, column=1, sticky="we", pady=4)

        ttk.Label(frame, text="密码").grid(row=3, column=0, sticky="e", padx=(0, 8), pady=4)
        self.password = tk.StringVar()
        entry = ttk.Entry(frame, textvariable=self.password, show="*", width=28)
        entry.grid(row=3, column=1, sticky="we", pady=4)

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
            entry.focus_set()
        else:
            self.username.focus_set()

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


class TableTab(ttk.Frame):
    """一个只读表格页签（领用表 / 计划表共用）。"""

    def __init__(self, parent, context, kind):
        ttk.Frame.__init__(self, parent, padding=8)
        self.context = context
        self.kind = kind  # requisitions / plans
        self.columns = REQUISITION_COLUMNS if kind == "requisitions" else PLAN_COLUMNS
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
        self.count_label = ttk.Label(toolbar, text="")
        self.count_label.pack(side="left", padx=12)

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)
        keys = [item[0] for item in self.columns]
        self.tree = ttk.Treeview(container, columns=keys, show="headings", height=16)
        for key, title, width in self.columns:
            self.tree.heading(key, text=title)
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
        messagebox.showinfo("提示", EDIT_HINT, parent=self.winfo_toplevel())

    def reload(self):
        keyword = (self.keyword.get() or "").strip()
        keys = [item[0] for item in self.columns]
        try:
            # 查询统一走仓储（数据层只有一处 SQL），界面只负责展示
            repository = (
                self.context.repositories.requisitions if self.kind == "requisitions" else self.context.repositories.plans
            )
            rows = repository.list(keyword=keyword or None, limit=500)
        except Exception as exc:
            messagebox.showerror("读取失败", str(exc), parent=self.winfo_toplevel())
            return

        self.tree.delete(*self.tree.get_children())
        for row in rows:
            values = []
            for key in keys:
                value = row[key]
                values.append("" if value is None else str(value))
            self.tree.insert("", "end", values=tuple(values))
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
    def __init__(self, parent, context, on_saved=None):
        self.context = context
        self.on_saved = on_saved
        self.window = tk.Toplevel(parent)
        self.window.title("邮件设置")
        self.window.geometry("620x420")
        self._build()

    def _build(self):
        settings = self.context.mail_settings
        frame = ttk.Frame(self.window, padding=12)
        frame.pack(fill="both", expand=True)

        info = ttk.LabelFrame(frame, text="当前配置（读自 app_settings）", padding=10)
        info.pack(fill="x")
        lines = [
            "启用：%s" % ("是" if settings.enabled else "否"),
            "服务器：%s:%d（安全方式 %s）" % (settings.host or "(未配置)", settings.port, settings.security),
            "发件人：%s <%s>" % (settings.from_name or "", settings.from_address or "(未配置)"),
            "账号：%s" % (settings.username or "(未配置)"),
            "密码来源：%s" % settings.password_source,
            "提醒：提前 %d 天，含逾期=%s，去重 %d 天" % (
                settings.warning_days_before,
                "是" if settings.warning_include_overdue else "否",
                settings.dedupe_days,
            ),
        ]
        for line in lines:
            ttk.Label(info, text=line).pack(anchor="w")

        local = ttk.LabelFrame(frame, text="本机 SMTP 口令（DPAPI 加密，只存本机）", padding=10)
        local.pack(fill="x", pady=10)
        ttk.Label(local, text="说明：共享库里的口令是主程序用 DPAPI(CurrentUser) 加密的，换机器解不开，"
                              "所以 XP 端需要在本机单独保存一份。").pack(anchor="w")
        self.password = tk.StringVar()
        ttk.Entry(local, textvariable=self.password, show="*", width=40).pack(anchor="w", pady=6)
        ttk.Button(local, text="保存到本机", command=self._save_password).pack(anchor="w")

        actions = ttk.Frame(frame)
        actions.pack(fill="x", pady=(4, 0))
        ttk.Button(actions, text="计划到期提醒（演练，不发信）", command=self._dry_run_reminder).pack(side="left")
        ttk.Button(actions, text="发送测试邮件（演练）", command=self._dry_run_test).pack(side="left", padx=8)

        self.output = tk.Text(frame, height=8, wrap="word")
        self.output.pack(fill="both", expand=True, pady=(10, 0))
        self._write("提示：真正发信需要 --send 或在后续里程碑的界面里显式确认。")

    def _write(self, text):
        self.output.configure(state="normal")
        self.output.insert("end", text + "\n")
        self.output.see("end")
        self.output.configure(state="disabled")

    def _save_password(self):
        password = self.password.get()
        if self.context.credentials.save_mail_password(password):
            self._write("本机口令已保存。" if password else "本机口令已清除。")
            if self.on_saved:
                self.on_saved()
        else:
            messagebox.showerror("保存失败", "DPAPI 不可用，无法加密保存。", parent=self.window)

    def _dry_run_reminder(self):
        from ..context import run_reminder

        self._write(run_reminder(self.context, dry_run=True))

    def _dry_run_test(self):
        recipient = self.context.mail_settings.from_address
        if not recipient:
            self._write("未配置发件地址，无法演练测试邮件。")
            return
        result = self.context.mail.send_test(recipient, dry_run=True)
        self._write("测试邮件演练：%s" % result.message)


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


def run(context):
    """启动界面；返回进程退出码。"""
    try:
        return Application(context).run()
    except tk.TclError as exc:
        compat.say_err("无法启动界面（tkinter/Tcl 错误）：%s" % exc)
        return 2
