using NLog;
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 高耗时任务协调器：一键报告生成、计划索引、报告扫描三者互斥。
    /// 用户启动新任务时，若已有任务在跑：
    /// 1. 调用其 <see cref="CancellationTokenSource.Cancel"/> 发停止信号；
    /// 2. 弹「正在中断」等待窗（非阻塞 Show + 内部 Wait 到当前任务真正结束）；
    /// 3. 关闭等待窗，再启动新任务。
    /// 任务实现方需在"当前文件夹/当前明细读完"后响应取消，保证数据完整。
    /// </summary>
    public class HighCostTaskCoordinator
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly object _gate = new();

        /// <summary>当前正在跑的任务名（null = 空闲）</summary>
        public string CurrentTaskName { get; private set; }

        /// <summary>当前任务的取消源（null = 空闲）</summary>
        public CancellationTokenSource CurrentCts { get; private set; }

        /// <summary>当前任务的完成 Task（用于等待其真正结束）</summary>
        private Task _currentTask;

        /// <summary>是否有任务正在跑</summary>
        public bool IsBusy
        {
            get { lock (_gate) return _currentTask != null && !_currentTask.IsCompleted; }
        }

        /// <summary>
        /// 请求启动一个高耗时任务：若已有任务在跑，先中断并等它结束，再启动新任务。
        /// </summary>
        /// <param name="taskName">任务名（用于中断提示与日志）</param>
        /// <param name="startAction">实际启动逻辑，参数是分配给该任务的 CancellationTokenSource</param>
        /// <returns>是否成功启动（被另一个未结束的任务卡住时返回 false）</returns>
        public async Task<bool> RequestStartAsync(string taskName, Func<CancellationTokenSource, Task> startAction)
        {
            if (startAction == null)
            {
                return false;
            }
            // 第一步：若有任务在跑，发取消信号并弹等待窗
            Task toWait = null;
            string interruptedName = null;
            lock (_gate)
            {
                if (_currentTask != null && !_currentTask.IsCompleted)
                {
                    toWait = _currentTask;
                    interruptedName = CurrentTaskName;
                    try
                    {
                        CurrentCts?.Cancel();
                    }
                    catch (Exception ex)
                    {
                        _logger.Warn($"发取消信号失败: {ex.Message}");
                    }
                }
            }
            if (toWait != null)
            {
                _logger.Info($"高耗时任务互斥：中断「{interruptedName}」以启动「{taskName}」");
                await WaitForInterruptAsync(interruptedName, toWait).ConfigureAwait(true);
            }

            // 第二步：登记新任务并启动
            CancellationTokenSource cts = new();
            Task run;
            lock (_gate)
            {
                if (_currentTask != null && !_currentTask.IsCompleted)
                {
                    // 极端情况下上一个任务还没结束（比如等待窗被绕过），拒绝启动避免并发
                    cts.Dispose();
                    _logger.Warn($"高耗时任务「{taskName}」启动被拒：上一个任务「{CurrentTaskName}」仍在跑");
                    return false;
                }
                CurrentTaskName = taskName;
                CurrentCts = cts;
                run = startAction(cts);
                _currentTask = run;
            }
            try
            {
                await run.ConfigureAwait(true);
            }
            catch (OperationCanceledException)
            {
                // 任务自己响应取消，正常结束
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"高耗时任务「{taskName}」执行异常");
            }
            finally
            {
                lock (_gate)
                {
                    if (ReferenceEquals(CurrentCts, cts))
                    {
                        CurrentTaskName = null;
                        CurrentCts = null;
                        _currentTask = null;
                    }
                }
                cts.Dispose();
            }
            return true;
        }

        /// <summary>
        /// 仅请求中断当前任务（不启动新任务），用于用户点「停止扫描」之类的入口。
        /// </summary>
        public void RequestInterruptCurrent()
        {
            lock (_gate)
            {
                if (_currentTask != null && !_currentTask.IsCompleted)
                {
                    try
                    {
                        CurrentCts?.Cancel();
                    }
                    catch (Exception ex)
                    {
                        _logger.Warn($"发取消信号失败: {ex.Message}");
                    }
                }
            }
        }

        /// <summary>
        /// 等待被中断的任务真正结束：期间弹一个非阻塞的「正在中断」等待窗，
        /// 任务结束后自动关闭。
        /// </summary>
        private static async Task WaitForInterruptAsync(string interruptedName, Task toWait)
        {
            Main.Views.WindowInterruptWait waitWindow = null;
            try
            {
                waitWindow = new Main.Views.WindowInterruptWait(interruptedName);
                waitWindow.Show();
            }
            catch (Exception)
            {
                // 等待窗打不开时退化为直接等（不阻塞 UI 线程：Task.WhenAny + Delay 让出）
                waitWindow = null;
            }
            try
            {
                await toWait.ConfigureAwait(true);
            }
            catch
            {
                // 被中断的任务可能抛 OperationCanceledException，属正常
            }
            finally
            {
                try
                {
                    waitWindow?.Close();
                }
                catch
                {
                }
            }
        }
    }
}
