using System;
using System.IO;

namespace ORT一键报告.Utils
{
    /// <summary>
    /// 数据文件夹不可用（路径不存在、没有权限、网络不可达）时抛出。
    /// 错误信息是给用户看的完整句子，界面直接显示。
    /// </summary>
    public class DataFolderUnavailableException : Exception
    {
        public DataFolderUnavailableException(string message) : base(message) { }
    }

    /// <summary>
    /// 文件夹工具：路径规范化、网络路径判断、可写性检测。
    /// </summary>
    public static class FolderUtil
    {
        /// <summary>
        /// 规范化文件夹路径（去引号、展开环境变量、相对路径按程序目录、去掉末尾分隔符）
        /// </summary>
        public static string Normalize(string folder)
        {
            if (string.IsNullOrWhiteSpace(folder))
            {
                return null;
            }
            string text = Environment.ExpandEnvironmentVariables(folder.Trim().Trim('"')).Trim();
            if (text.Length == 0)
            {
                return null;
            }
            if (!Path.IsPathRooted(text))
            {
                text = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, text);
            }
            try
            {
                text = Path.GetFullPath(text);
            }
            catch
            {
                return text;
            }
            // 去掉末尾分隔符（盘符根目录 "Z:\" 保留分隔符）
            while (text.Length > 3 &&
                   (text.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
                    || text.EndsWith(Path.AltDirectorySeparatorChar.ToString(), StringComparison.Ordinal)))
            {
                text = text.Substring(0, text.Length - 1);
            }
            return text;
        }

        /// <summary>
        /// 是否网络路径：UNC（\\服务器\共享）或映射的网络驱动器
        /// </summary>
        public static bool IsNetworkPath(string folder)
        {
            string path = Normalize(folder);
            if (string.IsNullOrEmpty(path))
            {
                return false;
            }
            if (path.StartsWith(@"\\", StringComparison.Ordinal) || path.StartsWith("//", StringComparison.Ordinal))
            {
                return true;
            }
            try
            {
                string root = Path.GetPathRoot(path);
                if (string.IsNullOrEmpty(root))
                {
                    return false;
                }
                return new DriveInfo(root).DriveType == DriveType.Network;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 确认文件夹存在并且可写（不存在则创建）；失败时返回 false 并给出给用户看的原因
        /// </summary>
        public static bool TryPrepare(string folder, out string error)
        {
            error = null;
            string path = Normalize(folder);
            if (string.IsNullOrEmpty(path))
            {
                error = "文件夹路径为空";
                return false;
            }
            try
            {
                Directory.CreateDirectory(path);
            }
            catch (Exception ex)
            {
                error = $"无法创建文件夹「{path}」：{ex.Message}";
                return false;
            }
            string probe = Path.Combine(path, ".ort_write_test_" + Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                File.WriteAllText(probe, "ok");
                File.Delete(probe);
                return true;
            }
            catch (Exception ex)
            {
                error = $"文件夹「{path}」不可写：{ex.Message}";
                return false;
            }
        }
    }
}
