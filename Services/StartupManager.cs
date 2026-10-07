using Microsoft.Win32;
using NLog;
using ORT一键报告.Utils;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 开机自启管理（当前用户）。启动项写在
    /// HKCU\Software\Microsoft\Windows\CurrentVersion\Run，命令行用**相对路径**：
    /// 相对路径由系统在当前用户登录时按该用户的程序目录解析（与资源管理器地址栏、Win+R 一致），
    /// 因此换了用户名也不必重写启动项。
    ///
    /// 远程路径（\\服务器\共享\... 、映射的网络盘符，以及 SUBST 出来的虚拟盘）里的程序
    /// **不做开机自启**：登录那一刻网络往往还没就绪、共享句柄也没有，写进去只会自启失败，
    /// 或者弹出找不到文件的错误。判定见 <see cref="TryDescribeRemoteLocation"/>。
    /// </summary>
    public static class StartupManager
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        /// <summary>注册表启动项路径（当前用户，无需管理员权限）</summary>
        private const string RunKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Run";

        /// <summary>启动项名称（任务管理器的「启动」页按这个名字显示）</summary>
        public const string ValueName = "ORT实验室管理系统";

        /// <summary>开机自启到后台时追加的命令行参数</summary>
        public const string BackgroundArgument = "--background";

        /// <summary>读不到注册表时的兜底错误说明</summary>
        private const string RegistryUnavailable = "无法访问注册表启动项（HKCU\\...\\Run）";

        /// <summary>
        /// 是否已在注册表里登记开机自启（只看注册表，不代表本次运行一定由自启拉起）
        /// </summary>
        public static bool IsEnabled()
        {
            try
            {
                using RegistryKey key = Registry.CurrentUser.OpenSubKey(RunKeyPath, false);
                return !string.IsNullOrWhiteSpace(key?.GetValue(ValueName) as string);
            }
            catch (Exception ex)
            {
                _logger.Warn($"读取开机自启设置失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 登记的开机自启命令行（没有登记时返回 null）
        /// </summary>
        public static string RegisteredCommandLine()
        {
            try
            {
                using RegistryKey key = Registry.CurrentUser.OpenSubKey(RunKeyPath, false);
                return key?.GetValue(ValueName) as string;
            }
            catch (Exception ex)
            {
                _logger.Warn($"读取开机自启命令行失败: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 登记的开机自启是否为「自启到后台」；没有登记时返回 false
        /// </summary>
        public static bool IsRegisteredToBackground()
            => ContainsBackgroundArgument(RegisteredCommandLine());

        /// <summary>
        /// 打开/关闭开机自启。返回 false 时 <paramref name="error"/> 说明原因（远程路径、写注册表失败等）。
        /// </summary>
        public static bool TryEnable(bool enabled, bool toBackground, out string error)
        {
            error = null;
            if (!enabled)
            {
                return TryDisable(out error);
            }
            string remote = TryDescribeRemoteLocation();
            if (remote != null)
            {
                error = remote;
                _logger.Warn($"拒绝登记开机自启：程序位于远程路径 {remote}");
                return false;
            }
            try
            {
                string commandLine = BuildCommandLine(toBackground);
                using RegistryKey key = Registry.CurrentUser.CreateSubKey(RunKeyPath, true);
                if (key == null)
                {
                    error = RegistryUnavailable;
                    return false;
                }
                key.SetValue(ValueName, commandLine, RegistryValueKind.String);
                _logger.Info($"已登记开机自启: {commandLine}");
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                _logger.Warn($"登记开机自启失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 取消开机自启（没登记也算成功）
        /// </summary>
        public static bool TryDisable(out string error)
        {
            error = null;
            try
            {
                using RegistryKey key = Registry.CurrentUser.OpenSubKey(RunKeyPath, true);
                if (key?.GetValue(ValueName) != null)
                {
                    key.DeleteValue(ValueName, false);
                    _logger.Info("已取消开机自启");
                }
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                _logger.Warn($"取消开机自启失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 开机自启命令行：相对路径 + 可选的后台参数
        /// </summary>
        public static string BuildCommandLine(bool toBackground)
        {
            string exe = RelativeExePath();
            return toBackground ? $"\"{exe}\" {BackgroundArgument}" : $"\"{exe}\"";
        }

        /// <summary>
        /// 程序 exe 的相对路径（相对当前用户的程序目录，如 Desktop\ORT\ORT实验室管理系统.exe）；
        /// 跨盘符等算不出相对路径时退回完整路径。
        /// </summary>
        public static string RelativeExePath()
        {
            string full = ProgramPath();
            if (string.IsNullOrEmpty(full))
            {
                return full;
            }
            try
            {
                // 远程路径本来就不允许自启，这里只保证算不出畸形字符串
                if (full.StartsWith(@"\\", StringComparison.Ordinal))
                {
                    return full;
                }
                // GetRelativePath 在跨盘符时原样返回绝对路径，这也是可以接受的写法
                return Report.GetRelativePath(
                    Environment.GetFolderPath(Environment.SpecialFolder.Programs), full);
            }
            catch (Exception ex)
            {
                _logger.Warn($"计算程序相对路径失败，改用完整路径: {ex.Message}");
                return full;
            }
        }

        /// <summary>
        /// 本次运行的程序完整路径（取不到时返回空串）
        /// </summary>
        public static string ProgramPath()
        {
            try
            {
                return Assembly.GetEntryAssembly()?.Location
                    ?? Process.GetCurrentProcess().MainModule?.FileName
                    ?? "";
            }
            catch (Exception ex)
            {
                _logger.Warn($"取程序路径失败: {ex.Message}");
                return "";
            }
        }

        /// <summary>
        /// 命令行是否要求「启动后直接进后台（托盘）」
        /// </summary>
        public static bool StartsInBackground(string[] args)
            => ContainsBackgroundArgument(args == null ? null : string.Join(" ", args));

        /// <summary>
        /// 命令行里是否带后台参数（只按空格分词，忽略大小写）
        /// </summary>
        public static bool ContainsBackgroundArgument(string commandLine)
        {
            if (string.IsNullOrWhiteSpace(commandLine))
            {
                return false;
            }
            string[] parts = commandLine.Split([' ', '\t'], StringSplitOptions.RemoveEmptyEntries);
            foreach (string part in parts)
            {
                if (string.Equals(part.Trim('"'), BackgroundArgument, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 程序是否位于不允许开机自启的远程位置（UNC / 网络盘 / SUBST 虚拟盘）
        /// </summary>
        public static bool IsRemoteLocation(string programPath = null)
            => TryDescribeRemoteLocation(programPath) != null;

        /// <summary>
        /// 远程位置说明（UNC 路径、网络盘符或 SUBST 盘符）；本机磁盘返回 null
        /// </summary>
        public static string TryDescribeRemoteLocation(string programPath = null)
        {
            string path = string.IsNullOrWhiteSpace(programPath) ? ProgramPath() : programPath;
            if (string.IsNullOrWhiteSpace(path))
            {
                return null;
            }
            if (path.StartsWith(@"\\", StringComparison.Ordinal)
                || path.StartsWith("//", StringComparison.Ordinal))
            {
                return path;
            }
            try
            {
                string full = Path.GetFullPath(path);
                string root = Path.GetPathRoot(full);
                if (string.IsNullOrEmpty(root) || root.Length < 2)
                {
                    return null;
                }
                // 映射的网络驱动器（DriveInfo 偶发取不到类型时用 Win32 再判一次）
                DriveType type = new DriveInfo(root).DriveType;
                if (type == DriveType.Network
                    || (type == DriveType.Unknown && GetDriveType(root) == DriveType.Network))
                {
                    return full;
                }
                // SUBST 出来的虚拟盘：登录时未必存在，等同远程处理
                if (type == DriveType.Fixed && IsSubstitutedDrive(root))
                {
                    return full;
                }
                return null;
            }
            catch (Exception ex)
            {
                _logger.Warn($"判断程序是否位于远程路径失败（{path}）: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 盘符是否由 SUBST 映射（如 subst X: C:\SomeDir）：
        /// 查该盘的 DOS 设备名，SUBST 会解析成 \??\X: 形式而不是 \Device\...。
        /// </summary>
        public static bool IsSubstitutedDrive(string root)
        {
            if (string.IsNullOrEmpty(root))
            {
                return false;
            }
            string letter = root.TrimEnd('\\', '/');
            if (letter.Length != 2 || letter[1] != ':')
            {
                return false;
            }
            StringBuilder target = new(512);
            try
            {
                if (QueryDosDevice(letter, target, target.Capacity) == 0)
                {
                    return false;
                }
                return target.ToString().StartsWith(@"\??\", StringComparison.Ordinal);
            }
            catch (Exception ex)
            {
                _logger.Warn($"判断 SUBST 盘符失败（{root}）: {ex.Message}");
                return false;
            }
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern uint QueryDosDevice(string deviceName, StringBuilder targetPath, int max);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        private static extern uint GetDriveTypeW(string rootPathName);

        private static DriveType GetDriveType(string root)
        {
            try
            {
                // Win32 返回值：2=可移动 3=固定 4=网络 5=光驱 6=内存盘
                return GetDriveTypeW(root) switch
                {
                    2 => DriveType.Removable,
                    3 => DriveType.Fixed,
                    4 => DriveType.Network,
                    5 => DriveType.CDRom,
                    6 => DriveType.Ram,
                    _ => DriveType.Unknown
                };
            }
            catch (Exception ex)
            {
                _logger.Warn($"读取驱动器类型失败（{root}）: {ex.Message}");
                return DriveType.Unknown;
            }
        }
    }
}
