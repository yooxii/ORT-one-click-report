using System;
using System.Collections.Generic;

namespace ORT一键报告.Models
{
    /// <summary>
    /// 本机设置（程序目录 Data\local_settings.json，**只有一个文件**）：
    /// 只保存"必须留在本机"的设置项——
    /// 1) 数据文件夹位置（数据库/附件/配图等全部数据都放在那里，可在远程共享上）；
    /// 2) ATE / EMI 源数据路径（各台电脑自己的目录）；
    /// 3)（旧位置遗留、仅用于迁移）登录 cookie 字段——凭据已改存到当前用户目录 %LocalAppData%（见 LoginCredentialStore）；
    /// 4) 计划表列布局。
    /// 其余设置（界面、邮件、业务路径等）随数据库放在数据文件夹里，多台电脑共用。
    /// 反序列化时缺失的项一律保持默认值，因此旧版本文件（或手工精简过的文件）都能读。
    /// </summary>
    public class LocalSettings
    {
        /// <summary>
        /// 数据文件夹：数据库、附件（OleFiles）、配图（PlanImages）等全部业务数据都在这里，
        /// 可以是本地目录，也可以是 \\服务器\共享\目录 之类的 UNC 路径；为空表示程序目录\Data。
        /// </summary>
        public string DataFolder { get; set; }

        /// <summary>ATE 源数据路径（本机目录）</summary>
        public string AteDataPath { get; set; }

        /// <summary>EMI 源数据路径（本机目录）</summary>
        public string EmiDataPath { get; set; }

        /// <summary>【旧位置遗留、仅用于迁移】登录用户名；凭据已改存用户目录，迁移后为 null</summary>
        public string LoginUsername { get; set; }

        /// <summary>【旧位置遗留、仅用于迁移】DPAPI 加密后的登录密码；凭据已改存用户目录，迁移后为 null</summary>
        public string LoginPasswordEnc { get; set; }

        /// <summary>【旧位置遗留、仅用于迁移】登录到期时间；凭据已改存用户目录，迁移后为 null</summary>
        public DateTime? LoginExpiry { get; set; }

        /// <summary>计划表列布局（"requisitions" / "plans" → 列键列表）</summary>
        public Dictionary<string, List<string>> PlansLayout { get; set; }

        /// <summary>后台运行设置（本机：最小化到后台 / 开机自启）</summary>
        public BackgroundSettings Background { get; set; } = new BackgroundSettings();
    }

    /// <summary>
    /// 后台运行设置（保存在本机 local_settings.json）：
    /// 「最小化到后台」只影响本机的窗口关闭行为，「开机自启」要写当前用户的注册表启动项，
    /// 因此这两项都跟随本机，不随数据库共享给其他电脑。
    /// 远程路径（网络共享 / 网络盘 / SUBST 虚拟盘）里的程序不登记开机自启，见 StartupManager。
    /// </summary>
    public class BackgroundSettings
    {
        /// <summary>关闭主窗口时询问是否最小化到后台（托盘），默认开启</summary>
        public bool MinimizeToTrayOnClose { get; set; } = true;

        /// <summary>开机自启（当前用户，无需管理员权限）</summary>
        public bool AutoStart { get; set; }

        /// <summary>开机自启时直接进后台（托盘），不弹出主窗口</summary>
        public bool AutoStartToBackground { get; set; }
    }
}
