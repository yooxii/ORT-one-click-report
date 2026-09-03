using System.Reflection;
using System.Resources;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 本地化辅助类，用于在代码中获取本地化字符串（如 MessageBox 消息）。
    /// 从 Resources.Strings 资源文件中读取。
    /// </summary>
    public static class LocalizationHelper
    {
        private static readonly ResourceManager _rm = new ResourceManager("ORT一键报告.Resources.Strings", Assembly.GetExecutingAssembly());

        /// <summary>
        /// 根据键名获取本地化字符串
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
