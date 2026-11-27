using Newtonsoft.Json;
using NLog;
using ORT一键报告.Models;
using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace ORT一键报告.Services.Sim44
{
    /// <summary>
    /// SIM44 数据库账号凭据的读写与加解密。
    /// 文件：%LocalAppData%\ORT实验室管理系统\sim44.json（与登录 cookie 同一个目录，互不覆盖）。
    /// 加密：DPAPI CurrentUser —— 换 Windows 用户、拷到别的电脑都解不开，正是我们要的。
    /// </summary>
    public static class Sim44CredentialStore
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private static readonly object Gate = new();

        /// <summary>公司要求：密码每 90 天变更一次</summary>
        public const int PasswordValidDays = 90;

        private static readonly string AppFolderName =
            Attribute.GetCustomAttribute(typeof(Sim44CredentialStore).Assembly,
                typeof(System.Reflection.AssemblyProductAttribute)) is System.Reflection.AssemblyProductAttribute p
                && !string.IsNullOrWhiteSpace(p.Product)
                ? p.Product
                : typeof(Sim44CredentialStore).Assembly.GetName().Name;

        /// <summary>凭据文件路径</summary>
        public static string FilePath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            AppFolderName, "sim44.json");

        /// <summary>读取凭据（不存在或读坏都返回 null）</summary>
        public static Sim44Credential Read()
        {
            lock (Gate)
            {
                try
                {
                    if (!File.Exists(FilePath))
                    {
                        return null;
                    }
                    string text = File.ReadAllText(FilePath);
                    return string.IsNullOrWhiteSpace(text)
                        ? null
                        : JsonConvert.DeserializeObject<Sim44Credential>(text);
                }
                catch (Exception ex)
                {
                    Logger.Warn($"读取 SIM44 凭据失败（按无凭据处理）：{ex.Message}");
                    return null;
                }
            }
        }

        /// <summary>保存凭据（密码在这里加密）</summary>
        public static void Save(string username, string plainPassword, DateTime? lastChangedAt, DateTime? lastLoginAt)
        {
            lock (Gate)
            {
                try
                {
                    Sim44Credential existing = ReadNoLock();
                    Sim44Credential cred = new()
                    {
                        Username = username,
                        PasswordEnc = string.IsNullOrEmpty(plainPassword) ? existing?.PasswordEnc : Encrypt(plainPassword),
                        LastChangedAt = lastChangedAt ?? existing?.LastChangedAt,
                        LastLoginAt = lastLoginAt ?? existing?.LastLoginAt,
                    };
                    WriteNoLock(cred);
                }
                catch (Exception ex)
                {
                    Logger.Error(ex, "保存 SIM44 凭据失败");
                }
            }
        }

        /// <summary>只更新「上次登录时间」（不动密码）</summary>
        public static void Touch(string username, DateTime lastLoginAt)
        {
            lock (Gate)
            {
                Sim44Credential cred = ReadNoLock();
                if (cred == null || !string.Equals(cred.Username, username, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }
                cred.LastLoginAt = lastLoginAt;
                WriteNoLock(cred);
            }
        }

        /// <summary>记录改密成功后的时间（90 天从这里重新算）</summary>
        public static void MarkPasswordChanged(string username, string newPassword, DateTime changedAt)
        {
            lock (Gate)
            {
                Sim44Credential cred = ReadNoLock() ?? new Sim44Credential();
                cred.Username = username;
                cred.PasswordEnc = Encrypt(newPassword);
                cred.LastChangedAt = changedAt;
                cred.LastLoginAt = DateTime.Now;
                WriteNoLock(cred);
            }
        }

        /// <summary>清除凭据（忘记密码/换账号时）</summary>
        public static void Clear()
        {
            lock (Gate)
            {
                try
                {
                    if (File.Exists(FilePath))
                    {
                        File.Delete(FilePath);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warn($"删除 SIM44 凭据失败：{ex.Message}");
                }
            }
        }

        /// <summary>解密密码；失败返回 null（换机器/换 Windows 用户就是这种情况）</summary>
        public static string DecryptPassword(string passwordEnc)
        {
            if (string.IsNullOrWhiteSpace(passwordEnc))
            {
                return null;
            }
            try
            {
                byte[] cipher = Convert.FromBase64String(passwordEnc);
                byte[] plain = ProtectedData.Unprotect(cipher, null, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(plain);
            }
            catch (Exception ex)
            {
                Logger.Warn($"解密 SIM44 密码失败（可能换了 Windows 用户或电脑）：{ex.Message}");
                return null;
            }
        }

        /// <summary>90 天到期的计算：返回剩余天数（负数表示已经超期）</summary>
        public static int? DaysUntilExpiry(DateTime? lastChangedAt)
        {
            if (lastChangedAt == null)
            {
                return null;
            }
            return PasswordValidDays - (int)(DateTime.Today - lastChangedAt.Value.Date).TotalDays;
        }

        private static string Encrypt(string plain)
        {
            byte[] plainBytes = Encoding.UTF8.GetBytes(plain ?? "");
            byte[] cipher = ProtectedData.Protect(plainBytes, null, DataProtectionScope.CurrentUser);
            return Convert.ToBase64String(cipher);
        }

        private static Sim44Credential ReadNoLock()
        {
            try
            {
                if (!File.Exists(FilePath))
                {
                    return null;
                }
                string text = File.ReadAllText(FilePath);
                return string.IsNullOrWhiteSpace(text) ? null : JsonConvert.DeserializeObject<Sim44Credential>(text);
            }
            catch
            {
                return null;
            }
        }

        private static void WriteNoLock(Sim44Credential cred)
        {
            string dir = Path.GetDirectoryName(FilePath);
            if (!string.IsNullOrEmpty(dir))
            {
                Directory.CreateDirectory(dir);
            }
            File.WriteAllText(FilePath, JsonConvert.SerializeObject(cred, Formatting.Indented));
        }
    }
}
