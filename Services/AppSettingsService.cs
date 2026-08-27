using Newtonsoft.Json;
using NLog;
using ORT一键报告.Models;
using System;
using System.IO;
using System.Windows;
using System.Windows.Media;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 应用设置服务：以 JSON 格式保存/读取设置（Data\settings.json）。
    /// 提供字体应用与路径默认目录查询；设置变更时触发 SettingsChanged。
    /// </summary>
    public class AppSettingsService
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly string _settingsFile;

        /// <summary>
        /// 当前设置（内存中）
        /// </summary>
        public AppSettings Settings { get; private set; } = new AppSettings();

        /// <summary>
        /// 设置变更事件（保存成功后触发）
        /// </summary>
        public event Action SettingsChanged;

        public AppSettingsService()
        {
            string dataDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data");
            Directory.CreateDirectory(dataDir);
            _settingsFile = Path.Combine(dataDir, "settings.json");
            Load();
        }

        /// <summary>
        /// 从 JSON 文件加载设置；文件不存在或损坏时使用默认值
        /// </summary>
        private void Load()
        {
            try
            {
                if (File.Exists(_settingsFile))
                {
                    Settings = JsonConvert.DeserializeObject<AppSettings>(File.ReadAllText(_settingsFile)) ?? new AppSettings();
                }
                else
                {
                    Settings = new AppSettings();
                }
                Settings.UI ??= new UiSettings();
                Settings.Paths ??= new PathSettings();
            }
            catch (Exception ex)
            {
                _logger.Warn($"设置文件损坏，使用默认设置: {ex.Message}");
                Settings = new AppSettings();
            }
        }

        /// <summary>
        /// 保存当前设置到 JSON 文件并通知变更
        /// </summary>
        public void Save()
        {
            try
            {
                File.WriteAllText(_settingsFile, JsonConvert.SerializeObject(Settings, Formatting.Indented));
                SettingsChanged?.Invoke();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "保存设置失败");
            }
        }

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
        public string AteDataDir => ValidDir(Settings.Paths.AteDataPath);
        public string EmiDataDir => ValidDir(Settings.Paths.EmiDataPath);
    }
}
