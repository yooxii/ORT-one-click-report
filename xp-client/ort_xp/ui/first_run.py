# -*- coding: utf-8 -*-
"""首次运行（或换了机器）：让用户挑出这个客户端要用的数据文件夹。

部署到 XP 机器上时，程序目录里通常**没有** ``ort_plans.db`` —— 数据库在主程序那边
（本机的 ``Data`` 目录，或映射成盘符的共享目录）。以前这种情况只写 stderr 就退出了，
窗口子系统下用户看到的就是「双击没反应」；现在改成：把原因讲清楚，并让用户直接选目录，
选完写进 ``<程序目录>\\Data\\local_settings.json``（与主程序同名同结构），下次启动直接生效。
"""

import os

#: 与主程序一致的数据库文件名
DB_FILE_NAME = "ort_plans.db"


def find_database_folder(folder):
    """在 ``folder`` 或它的 ``Data`` 子目录里找 ``ort_plans.db``，返回所在目录或 None。

    两种指法都常见：用户可能直接选中含库的 ``Data`` 目录，也可能只选到主程序目录
    （``bin\\Debug``）——后者再往下看一层 ``Data``。
    """
    if not folder:
        return None
    for candidate in (folder, os.path.join(folder, "Data")):
        if os.path.isfile(os.path.join(candidate, DB_FILE_NAME)):
            return candidate
    return None


def prompt_data_folder(context, reason, initial_dir=None):
    """弹出「选择数据文件夹」流程。

    返回值有三态，调用方按此区分处理：

    * 目录字符串 —— 用户选中了一个含 ``ort_plans.db`` 的文件夹；
    * ``""``     —— 用户已经看过说明并放弃（调用方不要再弹第二个框）；
    * ``None``   —— 界面不可用，压根没能弹出来（调用方改走日志 + 原生消息框）。
    """
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox
    except ImportError:
        return None

    try:
        root = tk.Tk()
    except Exception:
        return None
    root.withdraw()
    try:
        question = "%s\n\n要现在选择数据文件夹吗？\n（选错了可以取消，程序会退出并在日志里留下记录）" % reason
        if not messagebox.askyesno("没有找到数据库", question, parent=root):
            return ""

        directory = initial_dir
        if not directory:
            directory = context.data_dir if os.path.isdir(context.data_dir) else context.app_directory

        while True:
            options = {"parent": root, "title": "请选择含 %s 的数据文件夹" % DB_FILE_NAME, "initialdir": directory}
            try:
                chosen = filedialog.askdirectory(mustexist=True, **options)
            except TypeError:
                # 老版本 tk 不认 mustexist
                chosen = filedialog.askdirectory(**options)
            if not chosen:
                return ""
            found = find_database_folder(chosen)
            if found:
                return found
            messagebox.showwarning(
                "这个文件夹里没有 %s" % DB_FILE_NAME,
                "选中的文件夹：\n%s\n\n"
                "里面（以及它的 Data 子目录里）都没有 %s。\n"
                "数据库由主程序生成，请选中它的数据文件夹，或先用主程序把库建好。" % (chosen, DB_FILE_NAME),
                parent=root,
            )
            directory = chosen
    finally:
        try:
            root.destroy()
        except Exception:
            pass
