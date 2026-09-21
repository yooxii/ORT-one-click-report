using NLog;
using ORT一键报告.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 报告扫描进度快照（传给 UI 绑定用）
    /// </summary>
    public class ReportScanProgress
    {
        public int Processed { get; set; }
        public int Total { get; set; }
        public string CurrentFolder { get; set; }
        public bool IsRunning { get; set; }
        public bool WasInterrupted { get; set; }
    }

    /// <summary>
    /// 报告扫描服务：遍历报告根目录，按工作编号匹配报告夹，写 report_links 表；
    /// 匹配到的报告夹额外读 TestStatus 表判定报告状态（已完成/进行中），写回 plans.ReportStatus。
    /// 设计为可中断的高耗时任务：每处理完一个报告夹检查一次 CancellationToken，
    /// 当前报告夹的 TestStatus 读完再退出，保证数据完整。
    /// 与计划索引、一键报告生成通过 HighCostTaskCoordinator 互斥。
    /// </summary>
    public class ReportScanService
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;
        private readonly AppSettingsService _appSettings;
        private readonly ReportStatusReader _statusReader;

        public ReportScanService(DatabaseService db, AppSettingsService appSettings, ReportStatusReader statusReader)
        {
            _db = db;
            _appSettings = appSettings;
            _statusReader = statusReader;
        }

        /// <summary>进度/状态变化通知（已切回 UI 线程）</summary>
        public event Action Changed;

        /// <summary>扫描完成通知（参数为匹配到的报告夹数量；被中断时也会触发）</summary>
        public event Action<int> ScanCompleted;

        public bool IsRunning { get; private set; }
        public string CurrentFolder { get; private set; }
        public int Processed { get; private set; }
        public int Total { get; private set; }
        public bool WasInterrupted { get; private set; }

        /// <summary>
        /// 后台执行扫描。调用方通过 HighCostTaskCoordinator.RequestStartAsync 包装启动。
        /// </summary>
        public async Task RunAsync(CancellationToken ct)
        {
            if (IsRunning)
            {
                _logger.Info("报告扫描已在进行，跳过本次触发");
                return;
            }
            string root = _appSettings.ReportDir;
            if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root))
            {
                _logger.Info($"报告扫描跳过：报告路径未配置或不存在（{root}）");
                return;
            }
            List<string> jobNos = _db.FreeSql.Select<Plan>()
                .Where(p => p.JobNo != null && p.JobNo != "")
                .ToList(p => p.JobNo)
                .Where(j => !string.IsNullOrWhiteSpace(j))
                .Distinct()
                .ToList();
            if (jobNos.Count == 0)
            {
                _logger.Info("报告扫描跳过：计划表里没有工作编号");
                return;
            }

            IsRunning = true;
            WasInterrupted = false;
            Processed = 0;
            CurrentFolder = null;
            Total = CountCandidateDirs(root);
            RaiseChanged();

            List<ReportLink> found = [];
            Dictionary<string, string> statusUpdates = new();
            try
            {
                await Task.Run(() =>
                {
                    HashSet<string> matched = new();
                    foreach (string dir in EnumerateDirs(root, 4))
                    {
                        if (ct.IsCancellationRequested)
                        {
                            WasInterrupted = true;
                            break;
                        }
                        string name = Path.GetFileName(dir);
                        CurrentFolder = name;
                        string job = jobNos.FirstOrDefault(j => !matched.Contains(j)
                            && name.IndexOf(j, StringComparison.OrdinalIgnoreCase) >= 0);
                        if (job == null)
                        {
                            Processed++;
                            RaiseChanged();
                            continue;
                        }
                        string reportSub = null;
                        string overview = null;
                        try
                        {
                            reportSub = Directory.GetDirectories(dir)
                                .FirstOrDefault(d => Path.GetFileName(d).Equals("Report", StringComparison.OrdinalIgnoreCase));
                            overview = Directory.GetFiles(dir, "*.xls*")
                                .Where(f => !Path.GetFileName(f).StartsWith("~$", StringComparison.Ordinal))
                                .OrderByDescending(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                                .FirstOrDefault();
                        }
                        catch
                        {
                            Processed++;
                            RaiseChanged();
                            continue;
                        }
                        if (reportSub == null || overview == null)
                        {
                            Processed++;
                            RaiseChanged();
                            continue;
                        }
                        matched.Add(job);
                        found.Add(new ReportLink
                        {
                            JobNo = job,
                            ReportDir = reportSub,
                            OverviewFile = overview,
                            UpdatedAt = DateTime.Now
                        });

                        // 读 TestStatus 判定报告状态（读完当前文件夹再响应取消，保证数据完整）
                        string status = _statusReader.ReadStatus(overview);
                        if (status != null)
                        {
                            statusUpdates[job] = status;
                        }

                        Processed++;
                        RaiseChanged();
                        if (ct.IsCancellationRequested)
                        {
                            WasInterrupted = true;
                            break;
                        }
                    }
                }, ct).ConfigureAwait(true);
            }
            catch (OperationCanceledException)
            {
                WasInterrupted = true;
            }
            catch (Exception ex)
            {
                _logger.Warn($"报告扫描异常: {ex.Message}");
            }
            finally
            {
                try
                {
                    PersistResults(found, statusUpdates);
                }
                catch (Exception ex)
                {
                    _logger.Warn($"保存报告扫描结果失败: {ex.Message}");
                }
                IsRunning = false;
                CurrentFolder = null;
                RaiseChanged();
                _logger.Info($"报告扫描结束: 匹配 {found.Count} 个报告夹，状态更新 {statusUpdates.Count} 条{(WasInterrupted ? "（被中断）" : "")}");
                try
                {
                    ScanCompleted?.Invoke(found.Count);
                }
                catch (Exception ex)
                {
                    _logger.Warn($"报告扫描完成回调异常: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 把扫描结果写库：report_links 全量刷新；plans.ReportStatus 按 jobNo 增量更新，
        /// 且仅当当前值不是「无要求」时才覆盖。
        /// </summary>
        private void PersistResults(List<ReportLink> found, Dictionary<string, string> statusUpdates)
        {
            _db.FreeSql.Delete<ReportLink>().Where("1=1").ExecuteAffrows();
            if (found.Count > 0)
            {
                _db.FreeSql.Insert(found).ExecuteAffrows();
            }
            if (statusUpdates.Count == 0)
            {
                return;
            }
            foreach (KeyValuePair<string, string> kv in statusUpdates)
            {
                Plan plan = _db.FreeSql.Select<Plan>().Where(p => p.JobNo == kv.Key).First();
                if (plan == null)
                {
                    continue;
                }
                if (ReportStatusKind.IsUserLocked(plan.ReportStatus))
                {
                    continue;
                }
                if (plan.ReportStatus == kv.Value)
                {
                    continue;
                }
                plan.ReportStatus = kv.Value;
                plan.UpdatedAt = DateTime.Now;
                _db.FreeSql.Update<Plan>().SetSource(plan).Where(p => p.Id == plan.Id).ExecuteAffrows();
            }
        }

        private static int CountCandidateDirs(string root)
        {
            int count = 0;
            try
            {
                foreach (string _ in EnumerateDirs(root, 4))
                {
                    count++;
                }
            }
            catch
            {
            }
            return Math.Max(count, 1);
        }

        private static IEnumerable<string> EnumerateDirs(string root, int depth)
        {
            if (depth < 0)
            {
                yield break;
            }
            string[] dirs;
            try
            {
                dirs = Directory.GetDirectories(root);
            }
            catch
            {
                yield break;
            }
            foreach (string dir in dirs)
            {
                yield return dir;
                foreach (string sub in EnumerateDirs(dir, depth - 1))
                {
                    yield return sub;
                }
            }
        }

        private void RaiseChanged()
        {
            try
            {
                if (Application.Current?.Dispatcher != null && !Application.Current.Dispatcher.CheckAccess())
                {
                    Application.Current.Dispatcher.Invoke(() => Changed?.Invoke());
                }
                else
                {
                    Changed?.Invoke();
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"报告扫描进度通知失败: {ex.Message}");
            }
        }
    }
}
