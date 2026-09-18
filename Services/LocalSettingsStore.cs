using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NLog;
using ORT一键报告.Models;
using System;
using System.Collections.Generic;
using System.IO;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 本机设置文件（程序目录 Data\local_settings.json）的读写：
    /// 所有"必须留在本机"的设置项都存在这一个文件里（不再每个设置项单独建文件），
    /// 程序启动时读取，缺失项用默认值。
    /// 兼容迁移三种历史文件：
    /// - 旧版本地设置（字典格式，含 DatabasePath）；
    /// - 更早的 db_path.json；
    /// - 单独保存的 auth_cookie.json（登录 cookie）与 plans_layout.json（列布局）——
    ///   读入本文件后把旧文件删掉。
    /// </summary>
    public static class LocalSettingsStore
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private static readonly object Gate = new();

        /// <summary>本机设置文件（程序目录 Data 下，只有一个文件）</summary>
        public static string FilePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "local_settings.json");

        /// <summary>旧版数据库路径文件（历史遗留，迁移后删除）</summary>
        private static string LegacyDbPathFile => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "db_path.json");

        /// <summary>旧版本机登录 cookie 文件（迁移进本机设置后删除）</summary>
        private static string LegacyCookieFile => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "auth_cookie.json");

        /// <summary>旧版计划表列布局文件（迁移进本机设置后删除）</summary>
        private static string LegacyLayoutFile => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "plans_layout.json");

        /// <summary>
        /// 读取本机设置（缺失项用默认值）；读取失败返回默认设置
        /// </summary>
        public static LocalSettings Read()
        {
            lock (Gate)
            {
                LocalSettings settings = ReadFile(out bool legacyFormat);
                bool dirty = legacyFormat;
                dirty |= ImportLegacyCookie(settings);
                dirty |= ImportLegacyLayout(settings);
                if (dirty)
                {
                    WriteFile(settings);
                }
                return settings;
            }
        }

        /// <summary>
        /// 读取-修改-写回本机设置（加锁，原子替换；写失败只记日志）
        /// </summary>
        public static void Update(Action<LocalSettings> change)
        {
            if (change == null)
            {
                return;
            }
            lock (Gate)
            {
                LocalSettings settings = ReadFile(out _);
                ImportLegacyCookie(settings);
                ImportLegacyLayout(settings);
                change(settings);
                WriteFile(settings);
            }
        }

        /// <summary>
        /// 读取本机设置文件本体；legacyFormat 为 true 表示读到的是旧版字典格式（需要按新格式重写）
        /// </summary>
        private static LocalSettings ReadFile(out bool legacyFormat)
        {
            legacyFormat = false;
            try
            {
                string text = File.Exists(FilePath) ? File.ReadAllText(FilePath) : null;
                if (string.IsNullOrWhiteSpace(text) && File.Exists(LegacyDbPathFile))
                {
                    // 更早的版本：数据库路径单独一个文件
                    text = File.ReadAllText(LegacyDbPathFile);
                    legacyFormat = true;
                }
                if (string.IsNullOrWhiteSpace(text))
                {
                    return new LocalSettings();
                }
                JObject obj;
                try
                {
                    obj = JObject.Parse(text);
                }
                catch (JsonException)
                {
                    Logger.Warn("本机设置文件不是合法 JSON，已按默认设置处理");
                    return new LocalSettings();
                }
                // 旧格式：{ "DatabasePath": "...", "AteDataPath": "...", "EmiDataPath": "..." }
                //（DatabasePath 当时存的就是目录，现在改叫 DataFolder）
                if (obj["DataFolder"] == null && obj["DatabasePath"] != null)
                {
                    legacyFormat = true;
                    return new LocalSettings
                    {
                        DataFolder = (string)obj["DatabasePath"],
                        AteDataPath = (string)obj["AteDataPath"],
                        EmiDataPath = (string)obj["EmiDataPath"]
                    };
                }
                return obj.ToObject<LocalSettings>() ?? new LocalSettings();
            }
            catch (Exception ex)
            {
                Logger.Warn($"读取本机设置失败（按默认设置处理）: {ex.Message}");
                return new LocalSettings();
            }
        }

        private static void WriteFile(LocalSettings settings)
        {
            try
            {
                string dir = Path.GetDirectoryName(FilePath);
                if (!string.IsNullOrEmpty(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                File.WriteAllText(FilePath, JsonConvert.SerializeObject(settings, Formatting.Indented));
                // 迁移完成后清理旧文件
                DeleteQuietly(LegacyDbPathFile);
                DeleteQuietly(LegacyCookieFile);
                DeleteQuietly(LegacyLayoutFile);
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "保存本机设置失败");
            }
        }

        /// <summary>把旧版 auth_cookie.json 里的登录信息搬进本机设置</summary>
        private static bool ImportLegacyCookie(LocalSettings settings)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(settings.LoginUsername) && File.Exists(LegacyCookieFile))
                {
                    CookieShape cookie = JsonConvert.DeserializeObject<CookieShape>(File.ReadAllText(LegacyCookieFile));
                    if (cookie != null && !string.IsNullOrWhiteSpace(cookie.Username) && !string.IsNullOrWhiteSpace(cookie.PasswordEnc))
                    {
                        settings.LoginUsername = cookie.Username;
                        settings.LoginPasswordEnc = cookie.PasswordEnc;
                        settings.LoginExpiry = cookie.Expiry;
                        return true;
                    }
                    DeleteQuietly(LegacyCookieFile);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn($"迁移旧版本机登录信息失败: {ex.Message}");
            }
            return false;
        }

        /// <summary>把旧版 plans_layout.json 里的列布局搬进本机设置</summary>
        private static bool ImportLegacyLayout(LocalSettings settings)
        {
            try
            {
                if ((settings.PlansLayout == null || settings.PlansLayout.Count == 0) && File.Exists(LegacyLayoutFile))
                {
                    Dictionary<string, List<string>> layout = JsonConvert.DeserializeObject<Dictionary<string, List<string>>>(File.ReadAllText(LegacyLayoutFile));
                    if (layout != null && layout.Count > 0)
                    {
                        settings.PlansLayout = layout;
                        return true;
                    }
                    DeleteQuietly(LegacyLayoutFile);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn($"迁移旧版列布局失败: {ex.Message}");
            }
            return false;
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
                Logger.Warn($"删除旧设置文件失败（{path}）: {ex.Message}");
            }
        }

        /// <summary>旧版 auth_cookie.json 的结构（仅迁移用）</summary>
        private class CookieShape
        {
            public string Username { get; set; }
            public string PasswordEnc { get; set; }
            public DateTime Expiry { get; set; }
        }
    }
}
