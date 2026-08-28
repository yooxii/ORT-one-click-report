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
    /// 应用设置服务：常规设置项以键值对保存到数据库（app_settings 表）；
    /// 数据库路径/ATE数据路径/EMI数据路径保存在程序目录本地文件（local_settings.json）：
    /// 数据库路径因避免自引用必须独立于数据库，ATE/EMI 数据路径按需求与数据库路径同位置保存。
    /// 设置变更时触发 SettingsChanged。
    /// </summary>
    public class AppSettingsService
    {
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
        /// 从数据库加载设置；首次运行时兼容迁移旧版 settings.json 与旧版 ATE/EMI 数据库键
        /// </summary>
        private void Load()
        {
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
                        FontSize = values.TryGetValue("ui.fontSize", out string fs) && double.TryParse(fs, out double size) ? size : 14
                    },
                    Paths = new PathSettings
                    {
                        SchedulePath = values.TryGetValue("paths.schedule", out string v1) ? v1 : null,
                        RequisitionPath = values.TryGetValue("paths.requisition", out string v2) ? v2 : null,
                        ReportPath = values.TryGetValue("paths.report", out string v3) ? v3 : null
                    }
                };
            }
            catch (Exception ex)
            {
                _logger.Warn($"加载设置失败，使用默认设置: {ex.Message}");
                Settings = new AppSettings();
            }
        }

        /// <summary>
        /// 将数据库中旧版路径键迁移到本地设置文件，迁移后删除数据库键
        /// </summary>
        private void MigrateLocalKeys(Dictionary<string, string> values, string dbKey, string localKey)
        {
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
                    ["paths.schedule"] = Settings.Paths.SchedulePath,
                    ["paths.requisition"] = Settings.Paths.RequisitionPath,
                    ["paths.report"] = Settings.Paths.ReportPath
                };
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
        /// 将当前字体设置应用到指定窗口
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
            }
            catch (Exception ex)
            {
                _logger.Warn($"应用字体失败: {ex.Message}");
            }
        }

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
