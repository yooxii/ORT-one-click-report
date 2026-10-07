using NLog;
using System;
using System.IO;
using System.IO.Pipes;
using System.Threading;
using System.Threading.Tasks;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 单实例：同一个 Windows 用户只允许运行一个本程序。
    /// 第二个实例启动时通过命名管道通知第一个实例把主窗口显示出来（恢复/置前），然后自己退出，
    /// 因此重复双击图标只会「跳回已经打开的那个」，不会再开一个。
    ///
    /// 名字用 <c>Local\</c> 作用域：按登录会话隔离——同一台电脑的不同用户各自可以开一个，
    /// 与本程序「设置按本机、数据按文件夹」的定位一致。
    /// </summary>
    public static class SingleInstanceManager
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        /// <summary>互斥体名（Local\ 作用域 = 当前登录会话）</summary>
        public const string MutexName = @"Local\ORT实验室管理系统.SingleInstance";

        /// <summary>命名管道名（与互斥体同作用域）</summary>
        public const string PipeName = "ORT实验室管理系统.Activate";

        /// <summary>通知已有实例把窗口显示出来的管道消息</summary>
        public const string ActivateMessage = "ACTIVATE";

        /// <summary>管道读写超时（毫秒）：对端卡住时不至于把新实例一直挂在那里</summary>
        private const int PipeTimeoutMs = 3000;

        private static Mutex _mutex;
        private static volatile bool _listening;
        private static Action _activateHandler;

        /// <summary>
        /// 申请成为唯一实例。返回 true 表示本次是第一个实例，可以继续启动；
        /// 返回 false 表示已有实例在运行（已通知它显示窗口），调用方应立即退出。
        /// </summary>
        public static bool TryAcquire()
        {
            try
            {
                _mutex = new Mutex(true, MutexName, out bool createdNew);
                if (createdNew)
                {
                    _logger.Info("单实例检查通过（本次为第一个实例）");
                    return true;
                }
                _mutex.Dispose();
                _mutex = null;
                _logger.Info("检测到已有一个实例在运行，请求它显示窗口后退出本次启动");
                TrySignalExistingInstance();
                return false;
            }
            catch (Exception ex)
            {
                // 拿不到互斥体（权限异常等）不应该阻止程序启动
                _logger.Warn($"单实例检查失败，按可多开处理: {ex.Message}");
                return true;
            }
        }

        /// <summary>
        /// 第一个实例在窗口就绪后调用：开始监听「再开一个」的请求。
        /// <paramref name="onActivate"/> 在主线程（Dispatcher）之外被触发，由调用方自己切回 UI 线程。
        /// </summary>
        public static void StartListening(Action onActivate)
        {
            if (_listening || onActivate == null)
            {
                return;
            }
            _activateHandler = onActivate;
            _listening = true;
            _ = Task.Run(ListenLoop);
            _logger.Info("已开始监听重复启动请求");
        }

        /// <summary>
        /// 停止监听并释放互斥体（程序退出时调用；失败只记日志）
        /// </summary>
        public static void Release()
        {
            _listening = false;
            try
            {
                _mutex?.ReleaseMutex();
            }
            catch (Exception ex)
            {
                // 未持有互斥体时释放会抛异常，属正常情况
                _logger.Debug($"释放互斥体: {ex.Message}");
            }
            try
            {
                _mutex?.Dispose();
            }
            catch (Exception ex)
            {
                _logger.Warn($"释放单实例互斥体失败: {ex.Message}");
            }
            _mutex = null;
        }

        /// <summary>
        /// 通知已经运行的实例显示窗口（管道连不上说明对端正在退出，忽略即可）。
        /// 先快连两次（对端已在监听时不到 100ms 就返回，重复双击图标能立刻跳出窗口）；
        /// 再按较长间隔重试——应用内「重启」（如换数据文件夹）时新实例可能比旧实例先起来。
        /// </summary>
        private static void TrySignalExistingInstance()
        {
            for (int attempt = 1; attempt <= 40; attempt++)
            {
                int timeout = attempt <= 2 ? 200 : PipeTimeoutMs;
                try
                {
                    using NamedPipeClientStream client = new(".", PipeName, PipeDirection.Out, PipeOptions.None);
                    client.Connect(timeout);
                    using StreamWriter writer = new(client) { AutoFlush = true };
                    writer.Write(ActivateMessage);
                    return;
                }
                catch (Exception ex)
                {
                    _logger.Debug($"第 {attempt} 次通知已运行实例失败: {ex.Message}");
                    if (attempt < 3)
                    {
                        Thread.Sleep(100);
                    }
                    else
                    {
                        Thread.Sleep(250);
                    }
                }
            }
            _logger.Warn("无法通知已运行的实例显示窗口（它可能正在退出）");
        }

        /// <summary>
        /// 管道监听循环：一次只处理一个连接，处理完立刻重新监听。
        /// </summary>
        private static void ListenLoop()
        {
            while (_listening)
            {
                try
                {
                    using NamedPipeServerStream server = new(
                        PipeName, PipeDirection.In, 1, PipeTransmissionMode.Byte, PipeOptions.None);
                    server.WaitForConnection();
                    using StreamReader reader = new(server);
                    string message = reader.ReadLine();
                    if (!string.IsNullOrWhiteSpace(message)
                        && message.Trim().Equals(ActivateMessage, StringComparison.OrdinalIgnoreCase))
                    {
                        _activateHandler?.Invoke();
                    }
                }
                catch (Exception ex)
                {
                    if (!_listening)
                    {
                        break;
                    }
                    _logger.Warn($"处理重复启动请求失败: {ex.Message}");
                    Thread.Sleep(200);
                }
            }
        }
    }
}
