using System;

namespace ORT一键报告.Models
{
    /// <summary>
    /// SIM44（MES 回线工令）数据库账号凭据。
    /// 密码用 DPAPI（当前 Windows 用户）加密后存在用户目录，只有同一台电脑的同一个 Windows 用户能解；
    /// <see cref="LastChangedAt"/> 用来算「公司要求 90 天变更一次」的到期日。
    /// </summary>
    public class Sim44Credential
    {
        /// <summary>数据库账号</summary>
        public string Username { get; set; }

        /// <summary>DPAPI 加密后的密码（Base64）</summary>
        public string PasswordEnc { get; set; }

        /// <summary>上次修改密码的日期（用于 90 天到期计算；未知时为 null）</summary>
        public DateTime? LastChangedAt { get; set; }

        /// <summary>上次成功登录的时间（仅用于界面提示，不参与校验）</summary>
        public DateTime? LastLoginAt { get; set; }
    }
}
