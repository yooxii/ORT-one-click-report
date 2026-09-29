using System;

namespace ORT一键报告.Models
{
    /// <summary>
    /// 记住登录用的凭据，单独存放在【当前 Windows 用户】的本地应用数据目录
    /// （%LocalAppData%\ORT实验室管理系统\login.json），不再放在程序目录的本机设置里。
    /// 密码用 DPAPI（CurrentUser）加密——只有同一 Windows 用户能解，放用户目录才能同机多用户各自独立、互不覆盖。
    /// </summary>
    public class LoginCredential
    {
        /// <summary>用户名</summary>
        public string Username { get; set; }

        /// <summary>DPAPI（当前 Windows 用户）加密后的密码（Base64）</summary>
        public string PasswordEnc { get; set; }

        /// <summary>到期时间（过期后自动清除）</summary>
        public DateTime? Expiry { get; set; }
    }
}
