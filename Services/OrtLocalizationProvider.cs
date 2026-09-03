using System;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Reflection;
using System.Resources;
using System.Windows;
using WPFLocalizeExtension.Providers;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 自定义本地化提供程序。
    /// 直接使用已验证可用的 ResourceManager 读取 Resources.Strings 资源，
    /// 绕过 WPFLocalizeExtension 内置 ResxLocalizationProvider 对含中文程序集名
    /// （ORT一键报告）解析失败导致 UI 显示 "Key:xxx" 的问题。
    /// </summary>
    public class OrtLocalizationProvider : ILocalizationProvider
    {
        private const string ResourceBaseName = "ORT一键报告.Resources.Strings";

        private static readonly ResourceManager _rm =
            new ResourceManager(ResourceBaseName, Assembly.GetExecutingAssembly());

        /// <summary>
        /// 可用文化列表
        /// </summary>
        public ObservableCollection<CultureInfo> AvailableCultures { get; } =
            new ObservableCollection<CultureInfo>
            {
                new CultureInfo("zh-CN"),
                new CultureInfo("zh-TW"),
                new CultureInfo("en"),
            };

        public event ProviderChangedEventHandler ProviderChanged;
        public event ProviderErrorEventHandler ProviderError;
        public event ValueChangedEventHandler ValueChanged;

        /// <summary>
        /// 返回完全限定的资源键（用于引擎内部缓存与变更通知）
        /// </summary>
        public FullyQualifiedResourceKeyBase GetFullyQualifiedResourceKey(string key, DependencyObject target)
        {
            return new FQAssemblyDictionaryKey(key, "ORT一键报告", "Resources.Strings");
        }

        /// <summary>
        /// 获取本地化对象。优先按指定文化查找，找不到则回退到中性资源。
        /// </summary>
        public object GetLocalizedObject(string key, DependencyObject target, CultureInfo culture)
        {
            if (string.IsNullOrEmpty(key))
            {
                return null;
            }

            // 引擎传入的 key 为完全限定格式："程序集:字典:实际key"（两个冒号），取最后一个冒号后的实际 key
            int colonIdx = key.LastIndexOf(':');
            string actualKey = colonIdx >= 0 ? key.Substring(colonIdx + 1) : key;

            try
            {
                // 先尝试指定文化（zh-CN / zh-TW / en 及其卫星程序集）
                string value = culture != null ? _rm.GetString(actualKey, culture) : null;
                // 回退到中性资源（简体中文）
                if (value == null)
                {
                    value = _rm.GetString(actualKey, CultureInfo.InvariantCulture);
                }
                // 再回退到当前 UI 文化
                if (value == null)
                {
                    value = _rm.GetString(actualKey);
                }
                return value;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// 触发资源变更通知（切换语言时可调用以刷新 UI）
        /// </summary>
        public void RaiseProviderChanged()
        {
            ProviderChanged?.Invoke(this, new ProviderChangedEventArgs(null));
        }
    }
}
