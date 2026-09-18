using NLog;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace ORT一键报告.Utils
{
    /// <summary>
    /// UNC 路径（\\服务器\共享\目录）→ 网络驱动器盘符 的映射。
    ///
    /// 为什么需要：SQLite 的 Win32 层打不开 UNC 路径的库文件（无论是否 WAL、
    /// 是否加 \\?\UNC\ 前缀、是否指定 win32-longpath VFS，都会报
    /// "unable to open database file"），但**映射成盘符后完全正常**。
    /// 所以数据文件夹填 UNC 路径时，程序启动时自动把它映射成一个空闲盘符，
    /// 之后所有数据访问（数据库/附件/配图）都走盘符路径。
    /// </summary>
    public static class NetworkDriveMapper
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private const int RESOURCETYPE_DISK = 1;
        private const int ERROR_SUCCESS = 0;
        private const int ERROR_ALREADY_ASSIGNED = 85;
        private const int ERROR_SESSION_CREDENTIAL_CONFLICT = 1219;

        /// <summary>
        /// 是否 UNC 路径（\\服务器\共享...）
        /// </summary>
        public static bool IsUncPath(string path)
        {
            string text = FolderUtil.Normalize(path);
            return !string.IsNullOrEmpty(text)
                && (text.StartsWith(@"\\", StringComparison.Ordinal) || text.StartsWith("//", StringComparison.Ordinal));
        }

        /// <summary>
        /// 取 UNC 路径的共享根（\\服务器\共享）与共享内的相对部分（可能为空）
        /// </summary>
        public static bool TrySplitUnc(string uncPath, out string shareRoot, out string relative)
        {
            shareRoot = null;
            relative = null;
            string text = FolderUtil.Normalize(uncPath);
            if (string.IsNullOrEmpty(text) || !IsUncPath(text))
            {
                return false;
            }
            string normalized = text.Replace('/', '\\').TrimStart('\\');
            string[] parts = normalized.Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2)
            {
                return false;
            }
            shareRoot = @"\\" + parts[0] + "\\" + parts[1];
            relative = parts.Length > 2 ? string.Join("\\", parts, 2, parts.Length - 2) : "";
            return true;
        }

        /// <summary>
        /// 把 UNC 路径映射成网络驱动器并返回盘符形式的路径（已映射的共享会直接复用）
        /// </summary>
        public static bool TryMap(string uncPath, out string mappedPath, out string error)
        {
            mappedPath = null;
            error = null;
            try
            {
                if (!TrySplitUnc(uncPath, out string shareRoot, out string relative))
                {
                    error = "不是有效的 UNC 路径";
                    return false;
                }
                // 已经映射过同一个共享：直接复用那个盘符
                string existing = FindMappedDrive(shareRoot);
                if (existing != null)
                {
                    mappedPath = Combine(existing, relative);
                    Logger.Info($"数据文件夹复用已映射的网络驱动器：{shareRoot} → {existing}");
                    return true;
                }
                // 找一个空闲盘符（从 Z 往前找）
                foreach (char letter in FreeDriveLetters())
                {
                    int result = AddConnection(letter, shareRoot);
                    if (result == ERROR_SUCCESS)
                    {
                        mappedPath = Combine(letter + ":", relative);
                        Logger.Info($"数据文件夹已映射网络驱动器：{shareRoot} → {letter}:（{mappedPath}）");
                        return true;
                    }
                    if (result == ERROR_ALREADY_ASSIGNED)
                    {
                        continue; // 盘符被占：换一个
                    }
                    if (result == ERROR_SESSION_CREDENTIAL_CONFLICT)
                    {
                        error = $"已经用其他账号连接过 {shareRoot}（Windows 不允许同一台服务器使用两套凭据）。" +
                                "请先手工把该共享映射成一个盘符，或断开原有连接后重试。";
                        return false;
                    }
                    error = $"映射 {shareRoot} 失败：{DescribeWin32(result)}";
                    return false;
                }
                error = $"没有空闲的驱动器盘符可以映射 {shareRoot}，请手工断开一个网络驱动器后重试";
                return false;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        /// <summary>已映射到该共享的盘符（没有返回 null）</summary>
        private static string FindMappedDrive(string shareRoot)
        {
            for (char letter = 'Z'; letter >= 'D'; letter--)
            {
                string local = letter + ":";
                StringBuilder remote = new(512);
                int length = remote.Capacity;
                if (WNetGetConnection(local, remote, ref length) == ERROR_SUCCESS)
                {
                    string target = remote.ToString().TrimEnd('\\');
                    if (string.Equals(target, shareRoot, StringComparison.OrdinalIgnoreCase))
                    {
                        return local;
                    }
                }
            }
            return null;
        }

        /// <summary>当前未占用的盘符（从 Z 往前）</summary>
        private static IEnumerable<char> FreeDriveLetters()
        {
            HashSet<char> used = [];
            foreach (DriveInfo drive in DriveInfo.GetDrives())
            {
                if (drive.Name.Length > 0)
                {
                    used.Add(char.ToUpperInvariant(drive.Name[0]));
                }
            }
            for (char letter = 'Z'; letter >= 'D'; letter--)
            {
                if (!used.Contains(letter))
                {
                    yield return letter;
                }
            }
        }

        private static int AddConnection(char letter, string shareRoot)
        {
            NETRESOURCE resource = new()
            {
                dwType = RESOURCETYPE_DISK,
                lpLocalName = letter + ":",
                lpRemoteName = shareRoot
            };
            // 不传账号密码：沿用当前 Windows 用户已经建立的共享访问凭据；flags=0 表示本次登录会话内有效
            return WNetAddConnection2(ref resource, null, null, 0);
        }

        private static string Combine(string drive, string relative)
            => string.IsNullOrEmpty(relative) ? drive + "\\" : drive + "\\" + relative;

        private static string DescribeWin32(int code)
        {
            string message;
            try
            {
                message = new Win32Exception(code).Message;
            }
            catch
            {
                message = "未知错误";
            }
            return $"{message}（Win32 错误 {code}）";
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct NETRESOURCE
        {
            public int dwScope;
            public int dwType;
            public int dwDisplayType;
            public int dwUsage;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpLocalName;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpRemoteName;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpComment;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpProvider;
        }

        [DllImport("mpr.dll", CharSet = CharSet.Unicode, EntryPoint = "WNetAddConnection2W")]
        private static extern int WNetAddConnection2(ref NETRESOURCE netResource, string password, string username, int flags);

        [DllImport("mpr.dll", CharSet = CharSet.Unicode, EntryPoint = "WNetGetConnectionW")]
        private static extern int WNetGetConnection(string localName, StringBuilder remoteName, ref int length);
    }
}
