using Newtonsoft.Json;
using NLog;
using ORT一键报告.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Media;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 字重选项（设置界面下拉框绑定用）
    /// </summary>
    public class FontWeightOption
    {
        /// <summary>字重代码：Normal/Medium/SemiBold/Bold</summary>
        public string Code { get; }

        /// <summary>界面显示名（本地化）</summary>
        public string DisplayName { get; }

        public FontWeightOption(string code, string displayName)
        {
            Code = code;
            DisplayName = displayName;
        }
    }

    /// <summary>
    /// 应用设置服务：常规设置项以键值对保存到数据库（app_settings 表）；
    /// 数据库路径/ATE数据路径/EMI数据路径保存在程序目录本地文件（local_settings.json）：
    /// 数据库路径因避免自引用必须独立于数据库，ATE/EMI 数据路径按需求与数据库路径同位置保存。
    /// 设置变更时触发 SettingsChanged。
    /// </summary>
    public class AppSettingsService
    {
        /// <summary>
        /// 界面字号下限
        /// </summary>
        public const double MinUiFontSize = 8.0;

        /// <summary>
        /// 界面字号上限：再大即便按比例缩放，1300 像素宽的计划表也无法完整显示，
        /// 因此设置界面会警告并放弃超过该值的修改（见 WindowAppSettings.ApplyAll）。
        /// </summary>
        public const double MaxUiFontSize = 24.0;

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;

        /// <summary>
        /// 本地设置文件（程序目录 Data 下）：数据库路径/ATE数据路径/EMI数据路径
        /// </summary>
        public static string LocalSettingsFile => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "local_settings.json");

        /// <summary>
        /// 旧版本地设置文件名（仅数据库路径），用于一次性迁移
        /// </summary>
        private static string LegacyLocalFile => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "db_path.json");

        /// <summary>
        /// 本地设置缓存
        /// </summary>
        private Dictionary<string, string> _local = [];

        /// <summary>
        /// 当前设置（内存中，数据库部分）
        /// </summary>
        public AppSettings Settings { get; private set; } = new AppSettings();

        /// <summary>
        /// 设置变更事件（保存成功后触发）
        /// </summary>
        public event Action SettingsChanged;

        public AppSettingsService(DatabaseService db)
        {
            _db = db;
            LoadLocal();
            Load();
        }

        /* ###############################  本地设置文件（数据库路径/ATE/EMI）  ################################ */

        /// <summary>
        /// 加载本地设置文件；兼容迁移旧版 db_path.json
        /// </summary>
        private void LoadLocal()
        {
            try
            {
                if (File.Exists(LocalSettingsFile))
                {
                    _local = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText(LocalSettingsFile)) ?? [];
                }
                else if (File.Exists(LegacyLocalFile))
                {
                    // 一次性迁移旧文件
                    _local = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText(LegacyLocalFile)) ?? [];
                    SaveLocal();
                    File.Delete(LegacyLocalFile);
                }
                else
                {
                    _local = [];
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"读取本地设置文件失败: {ex.Message}");
                _local = [];
            }
        }

        private void SaveLocal()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LocalSettingsFile));
                File.WriteAllText(LocalSettingsFile, JsonConvert.SerializeObject(_local, Formatting.Indented));
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "保存本地设置文件失败");
            }
        }

        private string GetLocal(string key)
            => _local.TryGetValue(key, out string value) && !string.IsNullOrWhiteSpace(value) ? value : null;

        private void SetLocal(string key, string value)
        {
            _local[key] = value ?? "";
            SaveLocal();
        }

        /// <summary>
        /// 解析数据库文件路径：优先读取本地设置；未设置或无效时使用默认路径。
        /// 供 DatabaseService 初始化时调用（静态方法，避免依赖注入循环）。
        /// </summary>
        public static string ResolveDbPath()
        {
            string defaultPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "ort_plans.db");
            try
            {
                string file = File.Exists(LocalSettingsFile) ? LocalSettingsFile
                    : File.Exists(LegacyLocalFile) ? LegacyLocalFile : null;
                if (file != null)
                {
                    string dir = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText(file))?["DatabasePath"];
                    if (!string.IsNullOrWhiteSpace(dir))
                    {
                        Directory.CreateDirectory(dir);
                        return Path.Combine(dir, "ort_plans.db");
                    }
                }
            }
            catch
            {
                // 文件损坏时回退默认路径
            }
            return defaultPath;
        }

        public string GetDatabasePath() => GetLocal("DatabasePath");
        public string GetAteDataPath() => GetLocal("AteDataPath");
        public string GetEmiDataPath() => GetLocal("EmiDataPath");

        public void SetDatabasePath(string dir) => SetLocal("DatabasePath", dir);
        public void SetAteDataPath(string dir) => SetLocal("AteDataPath", dir);
        public void SetEmiDataPath(string dir) => SetLocal("EmiDataPath", dir);

        /* ###############################  数据库设置（app_settings 表）  ################################ */

        /// <summary>
        /// 读取布尔型设置（不存在时返回 default 值）
        /// </summary>
        public bool GetBool(string key, bool defaultValue)
        {
            string value = _db.FreeSql.Select<AppSetting>().Where(s => s.Key == key).First()?.Value;
            return bool.TryParse(value, out bool result) ? result : defaultValue;
        }

        /// <summary>
        /// 写入布尔型设置（存在则更新，不存在则插入）
        /// </summary>
        public void SetBool(string key, bool value)
        {
            AppSetting existing = _db.FreeSql.Select<AppSetting>().Where(s => s.Key == key).First();
            if (existing == null)
            {
                _db.FreeSql.Insert(new AppSetting { Key = key, Value = value.ToString() }).ExecuteAffrows();
            }
            else if (existing.Value != value.ToString())
            {
                existing.Value = value.ToString();
                _db.FreeSql.Update<AppSetting>().SetSource(existing).Where(s => s.Id == existing.Id).ExecuteAffrows();
            }
        }

        /// <summary>
        /// 读取整数型设置（不存在或无法解析时返回 defaultValue）
        /// </summary>
        public int GetInt(string key, int defaultValue)
        {
            string value = _db.FreeSql.Select<AppSetting>().Where(s => s.Key == key).First()?.Value;
            return int.TryParse(value, out int result) ? result : defaultValue;
        }

        /// <summary>
        /// 写入整数型设置（存在则更新，不存在则插入）
        /// </summary>
        public void SetInt(string key, int value) => SetText(key, value.ToString());

        /// <summary>
        /// 读取文本型设置（不存在时返回 defaultValue）
        /// </summary>
        public string GetText(string key, string defaultValue = null)
        {
            string value = _db.FreeSql.Select<AppSetting>().Where(s => s.Key == key).First()?.Value;
            return value ?? defaultValue;
        }

        /// <summary>
        /// 写入文本型设置（存在则更新，不存在则插入）
        /// </summary>
        public void SetText(string key, string value)
        {
            AppSetting existing = _db.FreeSql.Select<AppSetting>().Where(s => s.Key == key).First();
            if (existing == null)
            {
                _db.FreeSql.Insert(new AppSetting { Key = key, Value = value }).ExecuteAffrows();
            }
            else if (existing.Value != value)
            {
                existing.Value = value;
                _db.FreeSql.Update<AppSetting>().SetSource(existing).Where(s => s.Id == existing.Id).ExecuteAffrows();
            }
        }

        /// <summary>
        /// 从数据库加载设置；首次运行时兼容迁移旧版 settings.json 与旧版 ATE/EMI 数据库键
        /// </summary>
        private void Load()        {
            try
            {
                Dictionary<string, string> values = _db.FreeSql.Select<AppSetting>()
                    .ToDictionary(s => s.Key, s => s.Value);

                // 兼容迁移：旧版 JSON 设置文件
                if (values.Count == 0)
                {
                    MigrateFromLegacyJson(values);
                }

                // 兼容迁移：旧版保存在数据库中的 ATE/EMI 路径 → 本地文件
                MigrateLocalKeys(values, "paths.ate", "AteDataPath");
                MigrateLocalKeys(values, "paths.emi", "EmiDataPath");

                Settings = new AppSettings
                {
                    UI = new UiSettings
                    {
                        FontFamily = values.TryGetValue("ui.fontFamily", out string ff) && !string.IsNullOrWhiteSpace(ff) ? ff : "Microsoft YaHei UI",
                        FontSize = ResolveFontSize(values),
                        FontWeight = values.TryGetValue("ui.fontWeight", out string fw) && !string.IsNullOrWhiteSpace(fw) ? fw : "Normal"
                    },
                    Paths = new PathSettings
                    {
                        SchedulePath = values.TryGetValue("paths.schedule", out string v1) ? v1 : null,
                        RequisitionPath = values.TryGetValue("paths.requisition", out string v2) ? v2 : null,
                        ReportPath = values.TryGetValue("paths.report", out string v3) ? v3 : null
                    },
                    Mail = LoadMailSettings(values)
                };
            }
            catch (Exception ex)
            {
                _logger.Warn($"加载设置失败，使用默认设置: {ex.Message}");
                Settings = new AppSettings();
            }
        }

        /* ###############################  邮件设置  ################################ */

        /// <summary>
        /// 从键值对载入邮件设置（模板未自定义时回退到内置默认模板）
        /// </summary>
        private static MailSettings LoadMailSettings(Dictionary<string, string> values)
        {
            MailSettings mail = new()
            {
                Enabled = GetBool(values, "mail.enabled"),
                NoticeEnabled = GetBool(values, "mail.noticeEnabled", true),
                WarningEnabled = GetBool(values, "mail.warningEnabled", true),
                Host = GetString(values, "mail.host"),
                Port = GetInt(values, "mail.port", 25),
                Security = GetString(values, "mail.security") ?? "None",
                IgnoreCertErrors = GetBool(values, "mail.ignoreCertErrors"),
                UseDefaultCredentials = GetBool(values, "mail.useDefaultCredentials"),
                Username = GetString(values, "mail.username"),
                Password = Unprotect(GetString(values, "mail.passwordEnc")),
                FromAddress = GetString(values, "mail.fromAddress"),
                FromName = GetString(values, "mail.fromName"),
                CcList = GetString(values, "mail.ccList"),
                BodyIsHtml = GetBool(values, "mail.bodyIsHtml"),
                TimeoutSeconds = GetInt(values, "mail.timeoutSeconds", 30),
                WarningDaysBefore = GetInt(values, "mail.warningDaysBefore", 3),
                WarningIncludeOverdue = GetBool(values, "mail.warningIncludeOverdue", true),
                DedupeDays = GetInt(values, "mail.dedupeDays", 1)
            };
            foreach (MailTypeDefinition type in MailKind.All)
            {
                string subject = GetString(values, type.SubjectSettingKey);
                string body = GetString(values, type.BodySettingKey);
                if (!string.IsNullOrWhiteSpace(subject))
                {
                    mail.SetTemplate(type, true, subject);
                }
                if (!string.IsNullOrWhiteSpace(body))
                {
                    mail.SetTemplate(type, false, body);
                }
                // 抄送管理员默认：通知类开、警告类关（可在设置中按类型调整）
                mail.SetCcAdmins(type, GetBool(values, "mail.ccAdmin." + type.Code, type.Code == MailKind.Notice));
            }
            return mail;
        }

        private static string GetString(Dictionary<string, string> values, string key)
            => values.TryGetValue(key, out string value) && !string.IsNullOrWhiteSpace(value) ? value : null;

        private static bool GetBool(Dictionary<string, string> values, string key, bool fallback = false)
            => values.TryGetValue(key, out string value) && bool.TryParse(value, out bool result) ? result : fallback;

        private static int GetInt(Dictionary<string, string> values, string key, int fallback)
            => values.TryGetValue(key, out string value) && int.TryParse(value, out int result) ? result : fallback;

        /// <summary>
        /// 邮件相关键值对（密码 DPAPI 加密后写入密码键）
        /// </summary>
        private static void FillMailValues(Dictionary<string, string> values, MailSettings mail)
        {
            values["mail.enabled"] = mail.Enabled.ToString();
            values["mail.noticeEnabled"] = mail.NoticeEnabled.ToString();
            values["mail.warningEnabled"] = mail.WarningEnabled.ToString();
            values["mail.host"] = mail.Host;
            values["mail.port"] = mail.Port.ToString();
            values["mail.security"] = mail.Security ?? "None";
            values["mail.ignoreCertErrors"] = mail.IgnoreCertErrors.ToString();
            values["mail.useDefaultCredentials"] = mail.UseDefaultCredentials.ToString();
            values["mail.username"] = mail.Username;
            values["mail.passwordEnc"] = Protect(mail.Password);
            values["mail.fromAddress"] = mail.FromAddress;
            values["mail.fromName"] = mail.FromName;
            values["mail.ccList"] = mail.CcList;
            values["mail.bodyIsHtml"] = mail.BodyIsHtml.ToString();
            values["mail.timeoutSeconds"] = mail.TimeoutSeconds.ToString();
            values["mail.warningDaysBefore"] = mail.WarningDaysBefore.ToString();
            values["mail.warningIncludeOverdue"] = mail.WarningIncludeOverdue.ToString();
            values["mail.dedupeDays"] = mail.DedupeDays.ToString();
            foreach (MailTypeDefinition type in MailKind.All)
            {
                values[type.SubjectSettingKey] = mail.GetTemplate(type, true);
                values[type.BodySettingKey] = mail.GetTemplate(type, false);
                values["mail.ccAdmin." + type.Code] = mail.ShouldCcAdmins(type).ToString();
            }
        }

        /// <summary>
        /// DPAPI 加密（当前 Windows 用户），失败返回 null（不落库明文密码）
        /// </summary>
        private static string Protect(string plain)
        {
            if (string.IsNullOrEmpty(plain))
            {
                return null;
            }
            try
            {
                byte[] encrypted = System.Security.Cryptography.ProtectedData.Protect(
                    System.Text.Encoding.UTF8.GetBytes(plain), null,
                    System.Security.Cryptography.DataProtectionScope.CurrentUser);
                return Convert.ToBase64String(encrypted);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// DPAPI 解密（失败返回 null）
        /// </summary>
        private static string Unprotect(string encrypted)
        {
            if (string.IsNullOrEmpty(encrypted))
            {
                return null;
            }
            try
            {
                byte[] data = System.Security.Cryptography.ProtectedData.Unprotect(
                    Convert.FromBase64String(encrypted), null,
                    System.Security.Cryptography.DataProtectionScope.CurrentUser);
                return System.Text.Encoding.UTF8.GetString(data);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 将数据库中旧版路径键迁移到本地设置文件，迁移后删除数据库键
        /// </summary>
        private void MigrateLocalKeys(Dictionary<string, string> values, string dbKey, string localKey)        {
            if (values.TryGetValue(dbKey, out string value) && !string.IsNullOrWhiteSpace(value) && GetLocal(localKey) == null)
            {
                SetLocal(localKey, value);
            }
            if (values.ContainsKey(dbKey))
            {
                values.Remove(dbKey);
                _db.FreeSql.Delete<AppSetting>().Where(s => s.Key == dbKey).ExecuteAffrows();
            }
        }

        /// <summary>
        /// 迁移旧版 Data\settings.json（若存在），迁移成功后删除旧文件
        /// </summary>
        private void MigrateFromLegacyJson(Dictionary<string, string> values)
        {
            try
            {
                string legacyFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "settings.json");
                if (!File.Exists(legacyFile))
                {
                    return;
                }
                AppSettings legacy = JsonConvert.DeserializeObject<AppSettings>(File.ReadAllText(legacyFile));
                if (legacy == null)
                {
                    return;
                }
                values["ui.fontFamily"] = legacy.UI?.FontFamily;
                values["ui.fontSize"] = legacy.UI?.FontSize.ToString();
                values["ui.fontWeight"] = legacy.UI?.FontWeight;
                values["paths.schedule"] = legacy.Paths?.SchedulePath;
                values["paths.requisition"] = legacy.Paths?.RequisitionPath;
                values["paths.report"] = legacy.Paths?.ReportPath;
                SaveToDb(values);
                File.Delete(legacyFile);
                _logger.Info("已迁移旧版 settings.json 到数据库");
            }
            catch (Exception ex)
            {
                _logger.Warn($"迁移旧版设置文件失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 保存当前设置到数据库并通知变更
        /// </summary>
        public void Save()
        {
            try
            {
                Dictionary<string, string> values = new()
                {
                    ["ui.fontFamily"] = Settings.UI.FontFamily,
                    ["ui.fontSize"] = Settings.UI.FontSize.ToString(),
                    ["ui.fontWeight"] = Settings.UI.FontWeight,
                    ["paths.schedule"] = Settings.Paths.SchedulePath,
                    ["paths.requisition"] = Settings.Paths.RequisitionPath,
                    ["paths.report"] = Settings.Paths.ReportPath
                };
                FillMailValues(values, Settings.Mail);
                SaveToDb(values);
                SettingsChanged?.Invoke();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "保存设置失败");
            }
        }

        /// <summary>
        /// 键值对写入数据库（存在则更新，不存在则插入）
        /// </summary>
        private void SaveToDb(Dictionary<string, string> values)
        {
            foreach (KeyValuePair<string, string> kv in values)
            {
                AppSetting existing = _db.FreeSql.Select<AppSetting>().Where(s => s.Key == kv.Key).First();
                if (existing == null)
                {
                    _db.FreeSql.Insert(new AppSetting { Key = kv.Key, Value = kv.Value }).ExecuteAffrows();
                }
                else if (existing.Value != kv.Value)
                {
                    existing.Value = kv.Value;
                    _db.FreeSql.Update<AppSetting>().SetSource(existing).Where(s => s.Id == existing.Id).ExecuteAffrows();
                }
            }
        }

        /* ###############################  字体与目录  ################################ */

        /// <summary>
        /// 字号钳制到可用区间（历史数据可能存有超大字号，超出后界面无法完整显示）
        /// </summary>
        public static double ClampFontSize(double size)
            => size < MinUiFontSize ? MinUiFontSize : (size > MaxUiFontSize ? MaxUiFontSize : size);

        /// <summary>
        /// 读取已保存的字号并钳制到可用区间（超限时记录日志，按上限应用）
        /// </summary>
        private double ResolveFontSize(Dictionary<string, string> values)
        {
            if (!values.TryGetValue("ui.fontSize", out string raw) || !double.TryParse(raw, out double size))
            {
                return 14;
            }
            double clamped = ClampFontSize(size);
            if (Math.Abs(clamped - size) > 0.001)
            {
                _logger.Warn($"界面字号 {size} 超出 {MinUiFontSize}-{MaxUiFontSize} 范围，已按 {clamped} 应用");
            }
            return clamped;
        }

        /// <summary>
        /// 将当前字体设置应用到指定窗口（字号、字重、字体族，以及随字号放大的布局尺寸）
        /// </summary>
        public void ApplyFont(Window window)
        {
            if (window == null)
            {
                return;
            }
            try
            {
                FontFamily family = new(Settings.UI.FontFamily);
                window.FontFamily = family;
                window.FontSize = Settings.UI.FontSize;
                // 字重随设置逐级继承到未显式设置字重的控件（TextBox/Label 等隐式样式不再写死字重）
                window.FontWeight = ParseFontWeight(Settings.UI.FontWeight);
                // 布局随字号等比放大（等效 Windows 缩放）：窗口尺寸、表格列宽行高、Grid 绝对行列
                UiScale.Apply(window, UiScale.ScaleFor(Settings.UI.FontSize, MaxUiFontSize / UiScale.BaseFontSize));
            }
            catch (Exception ex)
            {
                _logger.Warn($"应用字体失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 字重代码 → WPF FontWeight（未知值按常规处理）
        /// </summary>
        public static FontWeight ParseFontWeight(string code) => code switch
        {
            "Medium" => FontWeights.Medium,
            "SemiBold" => FontWeights.SemiBold,
            "Bold" => FontWeights.Bold,
            _ => FontWeights.Normal
        };

        /// <summary>
        /// 将当前字体设置应用到所有已打开窗口
        /// </summary>
        public void ApplyFontToAll()
        {
            foreach (Window window in Application.Current.Windows)
            {
                ApplyFont(window);
            }
        }

        /// <summary>
        /// 获取有效目录（设置值为空或目录不存在时返回 null）
        /// </summary>
        private static string ValidDir(string path)
            => string.IsNullOrWhiteSpace(path) || !Directory.Exists(path) ? null : path;

        public string ScheduleDir => ValidDir(Settings.Paths.SchedulePath);
        public string RequisitionDir => ValidDir(Settings.Paths.RequisitionPath);
        public string ReportDir => ValidDir(Settings.Paths.ReportPath);
        public string AteDataDir => ValidDir(GetAteDataPath());
        public string EmiDataDir => ValidDir(GetEmiDataPath());
    }
}
