using System;
using System.Collections.Generic;

namespace ORT一键报告.Models
{
    /// <summary>
    /// 本机设置（程序目录 Data\local_settings.json，**只有一个文件**）：
    /// 只保存"必须留在本机"的设置项——
    /// 1) 数据文件夹位置（数据库/附件/配图等全部数据都放在那里，可在远程共享上）；
    /// 2) ATE / EMI 源数据路径（各台电脑自己的目录）；
    /// 3) 本机登录 cookie（用户名 + DPAPI 加密后的密码 + 到期时间）；
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

        /// <summary>本机登录 cookie：用户名</summary>
        public string LoginUsername { get; set; }

        /// <summary>本机登录 cookie：DPAPI（当前 Windows 用户）加密后的密码</summary>
        public string LoginPasswordEnc { get; set; }

        /// <summary>本机登录 cookie：到期时间</summary>
        public DateTime? LoginExpiry { get; set; }

        /// <summary>计划表列布局（"requisitions" / "plans" → 列键列表）</summary>
        public Dictionary<string, List<string>> PlansLayout { get; set; }
    }
}
