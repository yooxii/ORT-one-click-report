using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Markup;

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
        private const string CustomThemePathConfig = "current_theme_path.txt";
        public const string CustomCode = "Custom";
        private static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        /// <summary>当前已加载的主题字典（用于切换时移除）</summary>
        private static ResourceDictionary _currentThemeDict;

        /// <summary>
        /// 一套完整 UI 方案必须包含的语义化资源键（自定义主题需全部提供，否则校验失败）。
        /// 控件模板（ControlStyles.xaml）通过 DynamicResource 引用这些键实现换肤。
        /// </summary>
        public static readonly string[] RequiredKeys =
        {
            "WindowBackgroundBrush","CardBackgroundBrush","AlternateRowBrush","SidebarBackgroundBrush",
            "HeaderBackgroundBrush","TableHeaderForegroundBrush",
            "PrimaryBrush","PrimaryHoverBrush","PrimaryForegroundBrush",
            "ButtonBackgroundBrush","ButtonHoverBrush","ButtonPressedBrush","ButtonBorderBrush","ButtonForegroundBrush",
            "TextPrimaryBrush","TextSecondaryBrush","BorderSubtleBrush",
            "InputBackgroundBrush","InputForegroundBrush","InputBorderBrush","InputFocusBorderBrush",
            "TableRowHoverBrush","TableRowSelectedBrush","TableCellSelectedBrush",
            "ItemHoverBrush","ItemSelectedBrush",
            "StatusOkBrush","StatusOkBgBrush","StatusWarnBrush","StatusWarnBgBrush","StatusErrorBrush","StatusErrorBgBrush",
            "MenuBackgroundBrush","MenuForegroundBrush",
            "ScrollBarTrackBrush","ScrollBarThumbBrush","ScrollBarThumbHoverBrush",
            "FontFamilyUI","FontFamilyData",
            "CornerRadiusSmall","CornerRadiusButton","CornerRadiusCard","CornerRadiusLarge","DisabledOpacity",
        };

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
        /// 当前自定义主题文件路径（持久化）
        /// </summary>
        public static string CustomThemePath { get; private set; } =
            File.Exists(CustomThemePathConfig) ? File.ReadAllText(CustomThemePathConfig).Trim() : null;

        /// <summary>
        /// 初始化主题：优先读取用户上次保存的方案（含自定义），否则使用默认 Fluent
        /// </summary>
        public static void Initialize()
        {
            string saved = File.Exists(ThemeConfigPath)
                ? File.ReadAllText(ThemeConfigPath).Trim()
                : "Fluent";

            if (saved == CustomCode)
            {
                if (!string.IsNullOrEmpty(CustomThemePath) && File.Exists(CustomThemePath)
                    && TryLoadThemeDictionary(CustomThemePath, out ResourceDictionary dict, out _))
                {
                    ApplyDictionary(dict, CustomCode);
                    return;
                }
                logger.Warn("自定义主题文件缺失或无效，回退到 Fluent");
            }

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

        /// <summary>
        /// 从外部 XAML 文件加载主题字典并校验必需键（不应用，仅加载）。
        /// 供“导入自定义主题”使用。
        /// </summary>
        public static bool TryLoadThemeDictionary(string xamlPath, out ResourceDictionary dict, out string error)
        {
            dict = null;
            error = null;
            try
            {
                ResourceDictionary loaded;
                using (FileStream fs = File.OpenRead(xamlPath))
                {
                    loaded = (ResourceDictionary)XamlReader.Load(fs);
                }

                // out 参数不能在 lambda 中捕获，先用局部变量 loaded 校验
                var missing = RequiredKeys.Where(k => !loaded.Contains(k)).ToList();
                if (missing.Count > 0)
                {
                    error = "缺少必需资源键：" + string.Join("、", missing);
                    return false;
                }
                dict = loaded;
                return true;
            }
            catch (Exception ex)
            {
                error = "XAML 解析失败：" + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// 导入并应用自定义主题文件，成功后持久化为 Custom。
        /// 返回是否成功，失败时 error 为原因。
        /// </summary>
        public static bool ApplyCustomTheme(string xamlPath, out string error)
        {
            error = null;
            if (!TryLoadThemeDictionary(xamlPath, out ResourceDictionary dict, out error))
            {
                return false;
            }

            ApplyDictionary(dict, CustomCode);
            CustomThemePath = xamlPath;
            File.WriteAllText(CustomThemePathConfig, xamlPath);
            logger.Info("已导入自定义主题: {0}", xamlPath);
            return true;
        }

        /// <summary>
        /// 应用一个主题字典并持久化代码（内置/自定义共用）
        /// </summary>
        private static void ApplyDictionary(ResourceDictionary dict, string code)
        {
            if (_currentThemeDict != null)
            {
                Application.Current.Resources.MergedDictionaries.Remove(_currentThemeDict);
            }
            Application.Current.Resources.MergedDictionaries.Add(dict);
            _currentThemeDict = dict;
            CurrentTheme = code;
            File.WriteAllText(ThemeConfigPath, code);
            ThemeChanged?.Invoke();
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
