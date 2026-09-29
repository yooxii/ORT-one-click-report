using Newtonsoft.Json;
using NLog;
using ORT一键报告.Models;
using System;
using System.IO;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 登录凭据（记住登录）读写：单独存放在【当前 Windows 用户】的本地应用数据目录
    /// %LocalAppData%\ORT实验室管理系统\login.json。
    /// 放在用户目录而非程序目录，是因为密码用 DPAPI（CurrentUser）加密、只有同一 Windows 用户能解，
    /// 且同机多个 Windows 用户各自独立、互不覆盖；程序目录若在共享/只读位置也更稳妥。
    /// 首次读取时会自动把旧位置（程序目录 local_settings.json 的登录字段，或更早的 auth_cookie.json）迁移过来并清除。
    /// </summary>
    public static class LoginCredentialStore
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private static readonly object Gate = new();

        /// <summary>应用数据子目录名（取程序集产品名，与界面显示一致；产品名为空时回退到程序集名）</summary>
        private static readonly string AppFolderName =
            Attribute.GetCustomAttribute(typeof(LoginCredentialStore).Assembly, typeof(System.Reflection.AssemblyProductAttribute))
                is System.Reflection.AssemblyProductAttribute p && !string.IsNullOrWhiteSpace(p.Product)
                ? p.Product
                : typeof(LoginCredentialStore).Assembly.GetName().Name;

        /// <summary>登录凭据文件（当前 Windows 用户的本地应用数据目录下）</summary>
        public static string FilePath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            AppFolderName, "login.json");

        /// <summary>
        /// 读取登录凭据；不存在时尝试从旧位置迁移。都没有则返回 null。
        /// </summary>
        public static LoginCredential Read()
        {
            lock (Gate)
            {
                LoginCredential cred = ReadFile();
                if (cred != null && !string.IsNullOrWhiteSpace(cred.Username) && !string.IsNullOrWhiteSpace(cred.PasswordEnc))
                {
                    return cred;
                }
                return MigrateFromLegacy() ?? cred;
            }
        }

        /// <summary>保存登录凭据到用户目录</summary>
        public static void Save(string username, string passwordEnc, DateTime expiry)
        {
            lock (Gate)
            {
                WriteFile(new LoginCredential { Username = username, PasswordEnc = passwordEnc, Expiry = expiry });
            }
        }

        /// <summary>删除用户目录里的登录凭据文件（注销/过期时调用）</summary>
        public static void Clear()
        {
            lock (Gate)
            {
                DeleteQuietly(FilePath);
            }
        }

        private static LoginCredential ReadFile()
        {
            try
            {
                if (!File.Exists(FilePath))
                {
                    return null;
                }
                string text = File.ReadAllText(FilePath);
                return string.IsNullOrWhiteSpace(text) ? null : JsonConvert.DeserializeObject<LoginCredential>(text);
            }
            catch (Exception ex)
            {
                Logger.Warn($"读取登录凭据失败（按无凭据处理）: {ex.Message}");
                return null;
            }
        }

        private static void WriteFile(LoginCredential cred)
        {
            try
            {
                string dir = Path.GetDirectoryName(FilePath);
                if (!string.IsNullOrEmpty(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                File.WriteAllText(FilePath, JsonConvert.SerializeObject(cred, Formatting.Indented));
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "保存登录凭据失败");
            }
        }

        /// <summary>
        /// 从旧位置迁移登录凭据：旧版把凭据存在程序目录 local_settings.json（更早是 auth_cookie.json，
        /// 已由 LocalSettingsStore 读入 local_settings）。迁移到用户目录后清除旧位置字段，避免两处并存。
        /// </summary>
        private static LoginCredential MigrateFromLegacy()
        {
            try
            {
                LocalSettings local = LocalSettingsStore.Read();
                if (string.IsNullOrWhiteSpace(local.LoginUsername) || string.IsNullOrWhiteSpace(local.LoginPasswordEnc))
                {
                    return null;
                }
                LoginCredential cred = new()
                {
                    Username = local.LoginUsername,
                    PasswordEnc = local.LoginPasswordEnc,
                    Expiry = local.LoginExpiry
                };
                WriteFile(cred);
                // 清除程序目录里的旧凭据字段（迁移完成后不再使用）
                LocalSettingsStore.Update(s =>
                {
                    s.LoginUsername = null;
                    s.LoginPasswordEnc = null;
                    s.LoginExpiry = null;
                });
                Logger.Info("登录凭据已从程序目录迁移到用户目录");
                return cred;
            }
            catch (Exception ex)
            {
                Logger.Warn($"迁移旧登录凭据失败: {ex.Message}");
                return null;
            }
        }

        private static void DeleteQuietly(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn($"删除登录凭据文件失败（{path}）: {ex.Message}");
            }
        }
    }
}
