using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Resources;
using WPFLocalizeExtension.Engine;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 语言选项（用于 ComboBox 绑定）
    /// </summary>
    public class LanguageOption
    {
        public string Code { get; }
        public string DisplayName { get; }

        public LanguageOption(string code, string displayName)
        {
            Code = code;
            DisplayName = displayName;
        }
    }

    /// <summary>
    /// 语言本地化服务：初始化语言、切换语言、持久化保存、获取本地化字符串
    /// </summary>
    public static class LanguageService
    {
        private const string LangConfigPath = "current_language.txt";
        private static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        private static readonly ResourceManager _rm = new ResourceManager("ORT一键报告.Resources.Strings", Assembly.GetExecutingAssembly());

        /// <summary>
        /// 支持的语言列表
        /// </summary>
        public static LanguageOption[] SupportedLanguages { get; } =
        [
            new LanguageOption("zh-CN", "简体中文"),
            new LanguageOption("zh-TW", "繁體中文"),
            new LanguageOption("en", "English"),
        ];

        /// <summary>
        /// 初始化语言：优先读取用户上次保存的语言，否则使用系统语言
        /// </summary>
        public static void Initialize()
        {
            LocalizeDictionary.Instance.SetCurrentThreadCulture = true;

            string savedLang = File.Exists(LangConfigPath)
                ? File.ReadAllText(LangConfigPath).Trim()
                : CultureInfo.CurrentUICulture.Name;

            // 如果保存的语言不在支持列表中，使用简体中文
            bool supported = false;
            foreach (var lang in SupportedLanguages)
            {
                if (savedLang.StartsWith(lang.Code))
                {
                    savedLang = lang.Code;
                    supported = true;
                    break;
                }
            }
            if (!supported)
            {
                savedLang = "zh-CN";
            }

            SetLanguage(savedLang);
        }

        /// <summary>
        /// 语言变更事件（供 ViewModel 订阅以刷新非 XAML 绑定的本地化文本，如窗口标题）
        /// </summary>
        public static event Action LanguageChanged;

        /// <summary>
        /// 设置当前语言并持久化保存
        /// </summary>
        public static void SetLanguage(string cultureName)
        {
            try
            {
                LocalizeDictionary.Instance.Culture = new CultureInfo(cultureName);
                File.WriteAllText(LangConfigPath, cultureName);
                logger.Info("Language set to: {0}", cultureName);
                LanguageChanged?.Invoke();
            }
            catch (CultureNotFoundException)
            {
                logger.Error("Language {0} not found.", cultureName);
            }
        }

        /// <summary>
        /// 获取当前语言代码
        /// </summary>
        public static string GetCurrentLanguage()
        {
            return LocalizeDictionary.Instance.Culture.Name;
        }

        /// <summary>
        /// 获取本地化字符串（代码中使用，如 MessageBox）
        /// </summary>
        public static string Get(string key)
        {
            try
            {
                return _rm.GetString(key) ?? key;
            }
            catch
            {
                return key;
            }
        }
    }
}
