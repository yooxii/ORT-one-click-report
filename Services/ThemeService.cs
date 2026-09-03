using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;

namespace ORT一键报告.Services
{
    /// <summary>
    /// UI 主题服务：负责三套 UI 方案（Fluent / Material / DarkLab）的加载、切换与持久化。
    /// 主题以 ResourceDictionary 形式合并到 Application.Resources，
    /// 各窗口控件通过 DynamicResource 引用语义化资源键实现自动换肤。
    /// </summary>
    public static class ThemeService
    {
        private const string ThemeConfigPath = "current_theme.txt";
        private static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        /// <summary>当前已加载的主题字典（用于切换时移除）</summary>
        private static ResourceDictionary _currentThemeDict;

        /// <summary>
        /// 语言/主题变更事件（供 ViewModel 订阅刷新）
        /// </summary>
        public static event Action ThemeChanged;

        /// <summary>
        /// 支持的 UI 方案列表
        /// </summary>
        public static List<ThemeOption> SupportedThemes { get; } = new List<ThemeOption>
        {
            new ThemeOption("Fluent",  "Fluent 2 原生风格",  "Themes/FluentTheme.xaml"),
            new ThemeOption("Material", "Material 3 商务风", "Themes/MaterialTheme.xaml"),
            new ThemeOption("DarkLab", "深色实验室专业风",    "Themes/DarkLabTheme.xaml"),
        };

        /// <summary>
        /// 当前主题代码
        /// </summary>
        public static string CurrentTheme { get; private set; } = "Fluent";

        /// <summary>
        /// 初始化主题：优先读取用户上次保存的方案，否则使用默认 Fluent
        /// </summary>
        public static void Initialize()
        {
            string saved = File.Exists(ThemeConfigPath)
                ? File.ReadAllText(ThemeConfigPath).Trim()
                : "Fluent";

            bool found = false;
            foreach (ThemeOption opt in SupportedThemes)
            {
                if (opt.Code == saved)
                {
                    found = true;
                    break;
                }
            }
            ApplyTheme(found ? saved : "Fluent");
        }

        /// <summary>
        /// 应用指定主题并持久化保存
        /// </summary>
        public static void ApplyTheme(string code)
        {
            try
            {
                ThemeOption option = null;
                foreach (ThemeOption opt in SupportedThemes)
                {
                    if (opt.Code == code)
                    {
                        option = opt;
                        break;
                    }
                }
                if (option == null)
                {
                    logger.Warn("未知主题代码: {0}，回退到 Fluent", code);
                    option = SupportedThemes[0];
                }

                ResourceDictionary newDict = new ResourceDictionary
                {
                    Source = new Uri($"pack://application:,,,/{option.ResourcePath}", UriKind.Absolute)
                };

                // 移除旧主题，加入新主题（DynamicResource 引用会自动刷新）
                if (_currentThemeDict != null)
                {
                    Application.Current.Resources.MergedDictionaries.Remove(_currentThemeDict);
                }
                Application.Current.Resources.MergedDictionaries.Add(newDict);
                _currentThemeDict = newDict;
                CurrentTheme = option.Code;

                File.WriteAllText(ThemeConfigPath, option.Code);
                logger.Info("UI 主题已切换: {0}", option.Code);

                ThemeChanged?.Invoke();
            }
            catch (Exception ex)
            {
                logger.Error(ex, "应用主题 {0} 失败", code);
            }
        }
    }

    /// <summary>
    /// 主题选项（供 ComboBox 绑定）
    /// </summary>
    public class ThemeOption
    {
        public string Code { get; }
        public string DisplayName { get; }
        public string ResourcePath { get; }

        public ThemeOption(string code, string displayName, string resourcePath)
        {
            Code = code;
            DisplayName = displayName;
            ResourcePath = resourcePath;
        }
    }
}
