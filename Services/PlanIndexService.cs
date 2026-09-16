using FreeSql;
using NLog;
using ORT一键报告.Models;
using ORT一键报告.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 一次索引执行的结果
    /// </summary>
    public class PlanIndexRunResult
    {
        /// <summary>是否真的开始执行</summary>
        public bool Started { get; set; }

        /// <summary>说明（冲突/无数据/完成原因）</summary>
        public string Message { get; set; }

        /// <summary>本次处理条数</summary>
        public int Processed { get; set; }

        /// <summary>失败条数</summary>
        public int Failed { get; set; }

        /// <summary>是否已全部完成（含归并）</summary>
        public bool Completed { get; set; }
    }

    /// <summary>
    /// 归并统计
    /// </summary>
    public class PlanIndexMergeResult
    {
        public int TemplateCount { get; set; }
        public int PlanCount { get; set; }
        public int ItemCount { get; set; }
        public int DiffCount { get; set; }
        public int NewTestItems { get; set; }
    }

    /// <summary>
    /// 计划索引服务：扫描报告文件夹 → 逐份解析报告概览里的 ORT Plan 表 → 把测试项文本落库到原始表，
    /// 再统一归并出「测试项模板（重复内容只保存一次）+ 各机种计划的差异」。
    ///
    /// 任务与明细（plan_index_jobs / plan_index_entries）都落库，每条明细带认领信息，
    /// 因此同一数据库上的任一客户端都能接着未完成的明细继续跑（断点继续）；
    /// 多个客户端同时执行时通过认领时间做互斥，超时（默认 10 分钟）后才允许他人接管。
    /// </summary>
    public class PlanIndexService
    {
        /// <summary>认领超时（分钟）：超过后其他客户端可接管该任务/明细</summary>
        public const int ClaimTimeoutMinutes = 10;

        /// <summary>每批处理的明细条数（批间会刷新进度并让出线程）</summary>
        private const int BatchSize = 10;

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;
        private readonly AppSettingsService _settings;

        /// <summary>本客户端标识（多客户端互斥用）</summary>
        public static string ClientId { get; } = $"{Environment.MachineName}-{Environment.UserName}";

        private volatile bool _stopRequested;

        public PlanIndexService(DatabaseService db, AppSettingsService settings)
        {
            _db = db;
            _settings = settings;
        }

        /// <summary>进度/状态变化通知（已切回 UI 线程）</summary>
        public event Action Changed;

        /// <summary>是否正在本客户端执行</summary>
        public bool IsRunning { get; private set; }

        /// <summary>当前执行的任务Id</summary>
        public long CurrentJobId { get; private set; }

        /// <summary>最近一次状态说明</summary>
        public string StatusMessage { get; private set; } = "";

        /* ###############################  任务与明细  ################################ */

        /// <summary>
        /// 取某报告根目录下的最近一个索引任务（没有则返回 null）
        /// </summary>
        public PlanIndexJob GetLatestJob(string rootPath)
            => string.IsNullOrWhiteSpace(rootPath)
                ? null
                : _db.FreeSql.Select<PlanIndexJob>().Where(j => j.RootPath == rootPath).OrderByDescending(j => j.Id).First();

        /// <summary>全部索引任务（新的在前）</summary>
        public List<PlanIndexJob> GetJobs()
            => _db.FreeSql.Select<PlanIndexJob>().OrderByDescending(j => j.Id).ToList();

        /// <summary>某任务的明细</summary>
        public List<PlanIndexEntry> GetEntries(long jobId)
            => _db.FreeSql.Select<PlanIndexEntry>().Where(e => e.JobId == jobId).OrderBy(e => e.Id).ToList();

        /// <summary>
        /// 准备（或复用）索引任务：把报告根目录下新出现的报告夹补成待处理明细。
        /// </summary>
        /// <param name="rootPath">报告根目录</param>
        /// <param name="user">发起人</param>
        /// <param name="forceRebuild">true=全部重新索引（明细全部置回待处理并清空已抽取结果）</param>
        /// <param name="added">本次新增的明细数</param>
        public PlanIndexJob EnsureJob(string rootPath, string user, bool forceRebuild, out int added)
        {
            added = 0;
            if (string.IsNullOrWhiteSpace(rootPath) || !Directory.Exists(rootPath))
            {
                return null;
            }
            List<ReportFolder> folders = ScanReportFolders(rootPath);
            PlanIndexJob job = GetLatestJob(rootPath);
            if (job == null)
            {
                job = new PlanIndexJob
                {
                    Status = PlanIndexJob.StatusPending,
                    RootPath = rootPath,
                    StartedBy = user,
                    StartedAt = DateTime.Now,
                    UpdatedAt = DateTime.Now,
                    Message = "已建立索引任务"
                };
                job.Id = _db.FreeSql.Insert(job).ExecuteIdentity();
            }
            if (forceRebuild)
            {
                _db.FreeSql.Delete<PlanIndexEntry>().Where(e => e.JobId == job.Id).ExecuteAffrows();
                _db.FreeSql.Delete<PlanIndexRawItem>().Where(r => r.JobId == job.Id).ExecuteAffrows();
                job.Processed = 0;
                job.Failed = 0;
            }
            HashSet<string> known = _db.FreeSql.Select<PlanIndexEntry>()
                .Where(e => e.JobId == job.Id)
                .ToList()
                .Select(e => e.OverviewFile ?? "")
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (ReportFolder folder in folders)
            {
                if (!known.Add(folder.OverviewFile))
                {
                    continue;
                }
                _db.FreeSql.Insert(new PlanIndexEntry
                {
                    JobId = job.Id,
                    FolderName = folder.FolderName,
                    ModelName = folder.ModelName,
                    Stage = null,
                    OverviewFile = folder.OverviewFile,
                    Status = PlanIndexEntry.StatusPending,
                    UpdatedAt = DateTime.Now
                }).ExecuteAffrows();
                added++;
            }
            job.Total = (int)_db.FreeSql.Select<PlanIndexEntry>().Where(e => e.JobId == job.Id).Count();
            job.UpdatedAt = DateTime.Now;
            if (added > 0 && job.Status == PlanIndexJob.StatusDone)
            {
                job.Status = PlanIndexJob.StatusPending;
                job.Message = $"发现 {added} 份新报告，待继续索引";
            }
            _db.FreeSql.Update<PlanIndexJob>().SetSource(job).Where(j => j.Id == job.Id).ExecuteAffrows();
            return job;
        }

        /// <summary>
        /// 扫描报告根目录，找出所有"报告夹"：
        /// 文件夹里既有 Report 子目录、又有一个 Excel 概览文件（与计划表的扫描规则一致，
        /// 这样不会把 Report 子目录里单份试验报告当成一份报告夹）。
        /// </summary>
        public static List<ReportFolder> ScanReportFolders(string rootPath)
        {
            List<ReportFolder> found = [];
            if (string.IsNullOrWhiteSpace(rootPath) || !Directory.Exists(rootPath))
            {
                return found;
            }
            List<string> dirs = [rootPath];
            dirs.AddRange(EnumerateDirs(rootPath, 4));
            foreach (string dir in dirs)
            {
                if (!HasReportSubFolder(dir))
                {
                    continue;
                }
                string overview = FindOverviewFile(dir);
                if (overview == null)
                {
                    continue;
                }
                string name = Path.GetFileName(dir.TrimEnd(Path.DirectorySeparatorChar));
                found.Add(new ReportFolder
                {
                    FolderName = name,
                    ModelName = ParseModelName(name),
                    OverviewFile = overview
                });
            }
            return found;
        }

        /// <summary>目录下是否存在名为 Report 的子目录（报告夹的固定结构）</summary>
        private static bool HasReportSubFolder(string dir)
        {
            try
            {
                return Directory.GetDirectories(dir)
                    .Any(d => Path.GetFileName(d).Equals("Report", StringComparison.OrdinalIgnoreCase));
            }
            catch
            {
                return false;
            }
        }

        /// <summary>报告夹信息</summary>
        public class ReportFolder
        {
            /// <summary>文件夹名</summary>
            public string FolderName { get; set; }

            /// <summary>从文件夹名解析出的机种名称</summary>
            public string ModelName { get; set; }

            /// <summary>报告概览 Excel 路径</summary>
            public string OverviewFile { get; set; }
        }

        /// <summary>
        /// 目录下的报告概览文件：优先 .xlsx，跳过 Excel 临时文件（~$）
        /// </summary>
        private static string FindOverviewFile(string dir)
        {
            try
            {
                string[] files = Directory.GetFiles(dir, "*.xls*")
                    .Where(f => !Path.GetFileName(f).StartsWith("~$", StringComparison.Ordinal))
                    .OrderByDescending(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                    .ToArray();
                return files.FirstOrDefault();
            }
            catch
            {
                return null;
            }
        }

        /// <summary>从报告夹名称里取机种名（形如 FSA037-4B1G）</summary>
        public static string ParseModelName(string folderName)
        {
            Match match = Regex.Match(folderName ?? "", @"^([A-Za-z]{2,5}\d{2,5}-[A-Za-z0-9]+)");
            return match.Success ? match.Groups[1].Value.ToUpperInvariant() : null;
        }

        /// <summary>递归枚举子目录（限制深度）</summary>
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

        /* ###############################  执行（后台，可断点继续）  ################################ */

        /// <summary>请求停止（暂停在任务上，其他客户端/下次可继续）</summary>
        public void RequestStop() => _stopRequested = true;

        /// <summary>
        /// 后台执行索引：逐条处理待处理明细，全部完成后自动归并。
        /// 同一任务若已被其他客户端认领且未超时，则本次不重复执行。
        /// </summary>
        public async Task<PlanIndexRunResult> RunAsync(string rootPath, string user, bool forceRebuild = false)
        {
            PlanIndexRunResult result = new();
            if (IsRunning)
            {
                result.Message = "本客户端已在建立计划索引";
                return result;
            }
            if (string.IsNullOrWhiteSpace(rootPath) || !Directory.Exists(rootPath))
            {
                result.Message = "报告路径未配置或不存在，请先在设置里配置报告路径";
                return result;
            }
            PlanIndexJob job = EnsureJob(rootPath, user, forceRebuild, out int added);
            if (job == null)
            {
                result.Message = "没能建立索引任务";
                return result;
            }
            // 互斥：其他客户端正在执行（认领未超时）时不重复跑
            job = _db.FreeSql.Select<PlanIndexJob>().Where(j => j.Id == job.Id).First();
            if (job.ClaimedBy != null && job.ClaimedBy != ClientId
                && job.ClaimedAt.HasValue && job.ClaimedAt.Value > DateTime.Now.AddMinutes(-ClaimTimeoutMinutes))
            {
                result.Message = $"计划索引正在 {job.ClaimedBy} 上执行，稍后会自动继续";
                return result;
            }
            job.ClaimedBy = ClientId;
            job.ClaimedAt = DateTime.Now;
            job.Status = PlanIndexJob.StatusRunning;
            job.Message = "正在解析报告里的 ORT Plan";
            job.UpdatedAt = DateTime.Now;
            _db.FreeSql.Update<PlanIndexJob>().SetSource(job).Where(j => j.Id == job.Id).ExecuteAffrows();

            IsRunning = true;
            CurrentJobId = job.Id;
            _stopRequested = false;
            result.Started = true;
            StatusMessage = job.Message;
            RaiseChanged();

            int processed = 0, failed = 0;
            try
            {
                while (!_stopRequested)
                {
                    List<PlanIndexEntry> batch = ClaimNextBatch(job.Id);
                    if (batch.Count == 0)
                    {
                        break;
                    }
                    foreach (PlanIndexEntry entry in batch)
                    {
                        if (_stopRequested)
                        {
                            ReleaseEntry(entry);
                            break;
                        }
                        ProcessEntry(entry, job);
                        processed++;
                        if (entry.Status == PlanIndexEntry.StatusFailed)
                        {
                            failed++;
                        }
                        RefreshJobCounters(job);
                        RaiseChanged();
                    }
                    await Task.Delay(120).ConfigureAwait(false);
                }

                bool allSettled = _db.FreeSql.Select<PlanIndexEntry>()
                    .Where(e => e.JobId == job.Id && e.Status != PlanIndexEntry.StatusDone && e.Status != PlanIndexEntry.StatusFailed)
                    .Count() == 0;

                if (_stopRequested)
                {
                    job.Status = PlanIndexJob.StatusPaused;
                    job.Message = $"已暂停（已处理 {job.Processed}/{job.Total}），下次可继续";
                }
                else if (allSettled)
                {
                    job.Message = "正在归并测试项模板";
                    job.UpdatedAt = DateTime.Now;
                    _db.FreeSql.Update<PlanIndexJob>().SetSource(job).Where(j => j.Id == job.Id).ExecuteAffrows();
                    PlanIndexMergeResult merge = Merge(job.Id, user);
                    job.Status = PlanIndexJob.StatusDone;
                    job.FinishedAt = DateTime.Now;
                    job.Message = $"完成：测试项模板 {merge.TemplateCount} 个、机种计划 {merge.PlanCount} 个、"
                        + $"明细 {merge.ItemCount} 条、待确认差异 {merge.DiffCount} 处";
                    result.Completed = true;
                }
                else
                {
                    job.Status = PlanIndexJob.StatusPending;
                    job.Message = "暂未全部处理完，可继续执行";
                }
                result.Processed = processed;
                result.Failed = failed;
                result.Message = job.Message;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "建立计划索引失败");
                job.Status = PlanIndexJob.StatusFailed;
                job.Message = "失败：" + ex.Message;
                result.Message = job.Message;
            }
            finally
            {
                job.ClaimedBy = null;
                job.ClaimedAt = null;
                job.UpdatedAt = DateTime.Now;
                _db.FreeSql.Update<PlanIndexJob>().SetSource(job).Where(j => j.Id == job.Id).ExecuteAffrows();
                IsRunning = false;
                StatusMessage = job.Message;
                RaiseChanged();
            }
            return result;
        }

        /// <summary>
        /// 认领下一批待处理明细：
        /// 待处理，或被认领但已超时（客户端中途退出）的明细。
        /// </summary>
        private List<PlanIndexEntry> ClaimNextBatch(long jobId)
        {
            DateTime stale = DateTime.Now.AddMinutes(-ClaimTimeoutMinutes);
            List<PlanIndexEntry> batch = _db.FreeSql.Select<PlanIndexEntry>()
                .Where(e => e.JobId == jobId && (e.Status == PlanIndexEntry.StatusPending
                    || (e.Status == PlanIndexEntry.StatusRunning
                        && (e.ClaimedBy == null || e.ClaimedAt == null || e.ClaimedAt < stale))))
                .OrderBy(e => e.Id)
                .Limit(BatchSize)
                .ToList();
            foreach (PlanIndexEntry entry in batch)
            {
                entry.Status = PlanIndexEntry.StatusRunning;
                entry.ClaimedBy = ClientId;
                entry.ClaimedAt = DateTime.Now;
                entry.UpdatedAt = DateTime.Now;
                _db.FreeSql.Update<PlanIndexEntry>()
                    .Set(e => e.Status, entry.Status)
                    .Set(e => e.ClaimedBy, entry.ClaimedBy)
                    .Set(e => e.ClaimedAt, entry.ClaimedAt)
                    .Set(e => e.UpdatedAt, entry.UpdatedAt)
                    .Where(e => e.Id == entry.Id)
                    .ExecuteAffrows();
            }
            return batch;
        }

        /// <summary>释放认领（停止时把未处理的明细放回待处理）</summary>
        private void ReleaseEntry(PlanIndexEntry entry)
        {
            entry.Status = PlanIndexEntry.StatusPending;
            entry.ClaimedBy = null;
            entry.ClaimedAt = null;
            _db.FreeSql.Update<PlanIndexEntry>().SetSource(entry).Where(e => e.Id == entry.Id).ExecuteAffrows();
        }

        /// <summary>刷新任务的进度计数</summary>
        private void RefreshJobCounters(PlanIndexJob job)
        {
            job.Total = (int)_db.FreeSql.Select<PlanIndexEntry>().Where(e => e.JobId == job.Id).Count();
            job.Processed = (int)_db.FreeSql.Select<PlanIndexEntry>()
                .Where(e => e.JobId == job.Id && e.Status == PlanIndexEntry.StatusDone).Count();
            job.Failed = (int)_db.FreeSql.Select<PlanIndexEntry>()
                .Where(e => e.JobId == job.Id && e.Status == PlanIndexEntry.StatusFailed).Count();
            job.ClaimedAt = DateTime.Now; // 心跳，避免被其他客户端判为超时
            job.UpdatedAt = DateTime.Now;
            job.Message = $"正在解析报告（{job.Processed}/{job.Total}）";
            _db.FreeSql.Update<PlanIndexJob>().SetSource(job).Where(j => j.Id == job.Id).ExecuteAffrows();
        }

        /// <summary>
        /// 处理一条明细：解析报告概览里的 ORT Plan 表，把测试项原文写入原始表
        /// </summary>
        private void ProcessEntry(PlanIndexEntry entry, PlanIndexJob job)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(entry.OverviewFile) || !File.Exists(entry.OverviewFile))
                {
                    throw new FileNotFoundException("报告概览文件不存在", entry.OverviewFile);
                }
                NPOI.SS.UserModel.IWorkbook wb = ExcelNpoi.OpenAny(entry.OverviewFile);
                try
                {
                    ParsedOrtPlan plan = OrtPlanParser.ParseWorkbook(wb);
                    if (plan == null)
                    {
                        throw new InvalidOperationException("报告概览里没有 ORT Plan 工作表");
                    }
                    NPOI.SS.UserModel.ISheet cover = FindCoverSheet(wb);
                    string model = FirstNonEmpty(Report.FindInfoByText(cover, "Model Name")?.Data, entry.ModelName,
                        ParseModelName(entry.FolderName));
                    string stage = PlanStage.Normalize(Report.FindInfoByText(cover, "Product Stage")?.Data);
                    string note = plan.Note;

                    _db.FreeSql.Delete<PlanIndexRawItem>().Where(r => r.EntryId == entry.Id).ExecuteAffrows();
                    int order = 0;
                    List<PlanIndexRawItem> raws = [];
                    foreach (ParsedOrtPlanRow row in plan.Items)
                    {
                        order++;
                        raws.Add(new PlanIndexRawItem
                        {
                            JobId = job.Id,
                            EntryId = entry.Id,
                            ModelName = model,
                            Stage = stage,
                            Category = row.Category,
                            TestItemName = row.TestItemName,
                            OrderNo = order,
                            SamplingPlan = row.SamplingPlan,
                            TestCondition = row.TestCondition,
                            PassCriterion = row.PassCriterion,
                            Remark = row.Remark
                        });
                    }
                    if (raws.Count > 0)
                    {
                        _db.FreeSql.Insert(raws).ExecuteAffrows();
                    }
                    // 顺手把 ORT Plan 表里的图片按测试项抽出来（生成新报告模板时一并写入）
                    IndexOrtPlanImages(wb, plan, entry.OverviewFile, model);
                    entry.ModelName = model;
                    entry.Stage = stage;
                    entry.Note = Truncate(note, 1000);
                    entry.ItemCount = raws.Count;
                    entry.Status = PlanIndexEntry.StatusDone;
                    entry.Error = null;
                }
                finally
                {
                    wb.Close();
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"索引报告失败（{entry.FolderName}）：{ex.Message}");
                entry.Status = PlanIndexEntry.StatusFailed;
                entry.Error = Truncate(ex.Message, 500);
            }
            finally
            {
                entry.ClaimedBy = null;
                entry.ClaimedAt = null;
                entry.UpdatedAt = DateTime.Now;
                _db.FreeSql.Update<PlanIndexEntry>().SetSource(entry).Where(e => e.Id == entry.Id).ExecuteAffrows();
            }
        }

        /// <summary>
        /// 从历史报告的 ORT Plan 表里抽出配图，按锚点所在行归到对应测试项目上。
        /// 图片落盘到数据库目录下的 PlanImages，数据库里只记相对文件名；
        /// 同一个测试项目以最新一份报告为准（覆盖式更新，重建索引不会越积越多）。
        /// </summary>
        private void IndexOrtPlanImages(NPOI.SS.UserModel.IWorkbook wb, ParsedOrtPlan plan, string overviewFile, string model)
        {
            try
            {
                NPOI.SS.UserModel.ISheet sheet = OrtPlanParser.FindSheet(wb);
                List<ExcelNpoi.SheetPicture> pictures = ExcelNpoi.PictureDetails(sheet);
                List<ParsedOrtPlanRow> itemRows = plan.Items.Where(r => r.Row > 0 && !string.IsNullOrWhiteSpace(r.TestItemName)).ToList();
                if (pictures.Count == 0 || itemRows.Count == 0)
                {
                    return;
                }
                // 1. 图片 → 测试项：取锚点行（左上角）之前最近的一个测试项；在表头之前的一律忽略（表头只有 logo）
                Dictionary<string, List<(string Name, string Key, byte[] Bytes)>> matched = [];
                foreach (ExcelNpoi.SheetPicture picture in pictures)
                {
                    if (picture.Bytes == null || picture.Bytes.Length == 0 || picture.Row <= plan.HeaderRow)
                    {
                        continue; // 表头区的公司 logo 不算测试项配图
                    }
                    ParsedOrtPlanRow row = itemRows.LastOrDefault(r => r.Row <= picture.Row);
                    if (row == null)
                    {
                        continue;
                    }
                    string key = NameKey(row.TestItemName);
                    if (string.IsNullOrEmpty(key))
                    {
                        continue;
                    }
                    if (!matched.TryGetValue(key, out List<(string, string, byte[])> list))
                    {
                        list = [];
                        matched[key] = list;
                    }
                    list.Add((row.TestItemName, key, picture.Bytes));
                }
                if (matched.Count == 0)
                {
                    return;
                }

                // 2. 覆盖式落库：先清掉这些测试项的旧记录与旧文件
                Directory.CreateDirectory(_db.PlanImagesDir);
                foreach (KeyValuePair<string, List<(string Name, string Key, byte[] Bytes)>> pair in matched)
                {
                    foreach (PlanItemImage old in _db.FreeSql.Select<PlanItemImage>().Where(i => i.NameKey == pair.Key).ToList())
                    {
                        DeleteImageFile(old.FileName);
                    }
                    _db.FreeSql.Delete<PlanItemImage>().Where(i => i.NameKey == pair.Key).ExecuteAffrows();
                    int order = 0;
                    string safeKey = SafeFileName(pair.Key);
                    foreach ((string name, string key, byte[] bytes) in pair.Value)
                    {
                        order++;
                        string fileName = $"{safeKey}_{order}.png";
                        string fullPath = Path.Combine(_db.PlanImagesDir, fileName);
                        File.WriteAllBytes(fullPath, bytes);
                        GetImageSize(bytes, out int width, out int height);
                        _db.FreeSql.Insert(new PlanItemImage
                        {
                            NameKey = key,
                            TestItemName = name,
                            ModelName = model,
                            SourceFile = Truncate(overviewFile, 500),
                            FileName = fileName,
                            WidthPx = width,
                            HeightPx = height,
                            OrderNo = order,
                            UpdatedAt = DateTime.Now
                        }).ExecuteAffrows();
                    }
                }
                _logger.Info($"已抽取 ORT Plan 配图：{matched.Count} 个测试项目（{Path.GetFileName(overviewFile)}）");
            }
            catch (Exception ex)
            {
                _logger.Warn($"抽取 ORT Plan 配图失败（{Path.GetFileName(overviewFile)}）：{ex.Message}");
            }
        }

        /// <summary>删掉旧的配图文件（文件不存在当作已删）</summary>
        private void DeleteImageFile(string fileName)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(fileName))
                {
                    return;
                }
                string path = Path.Combine(_db.PlanImagesDir, fileName);
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"删除旧配图失败（{fileName}）：{ex.Message}");
            }
        }

        /// <summary>图片像素尺寸（读不出来时按 0 记）</summary>
        private static void GetImageSize(byte[] bytes, out int width, out int height)
        {
            width = 0;
            height = 0;
            try
            {
                using MemoryStream stream = new(bytes);
                using System.Drawing.Image image = System.Drawing.Image.FromStream(stream);
                width = image.Width;
                height = image.Height;
            }
            catch (Exception)
            {
                // 认不出尺寸就算了，生成报告时按默认大小放
            }
        }

        /// <summary>归一化键转成安全的文件名</summary>
        private static string SafeFileName(string key)
        {
            StringBuilder builder = new();
            foreach (char ch in key ?? "")
            {
                if (char.IsLetterOrDigit(ch))
                {
                    builder.Append(ch);
                }
            }
            return builder.Length == 0 ? "item" : builder.ToString();
        }

        /// <summary>取 Cover 工作表（按名称找，找不到退回第一张表）</summary>
        private static NPOI.SS.UserModel.ISheet FindCoverSheet(NPOI.SS.UserModel.IWorkbook wb)
        {
            for (int i = 0; i < wb.NumberOfSheets; i++)
            {
                string name = wb.GetSheetName(i);
                if (!string.IsNullOrWhiteSpace(name) && name.ToLowerInvariant().Contains("cover"))
                {
                    return wb.GetSheetAt(i);
                }
            }
            return ExcelNpoi.SheetAt(wb, 0);
        }

        private static string FirstNonEmpty(params string[] values)
        {
            foreach (string value in values)
            {
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value.Trim();
                }
            }
            return null;
        }

        private static string Truncate(string text, int max)
            => string.IsNullOrEmpty(text) || text.Length <= max ? text : text.Substring(0, max);

        /* ###############################  归并（压缩重复内容）  ################################ */

        /// <summary>
        /// 归并：把原始抽取结果压缩成「测试项模板 + 各机种计划的差异」。
        /// 同一测试项目的每段文本取出现次数最多的写法作为模板，其余按机种存为差异并列入待确认。
        /// </summary>
        public PlanIndexMergeResult Merge(long jobId, string user)
        {
            PlanIndexMergeResult result = new();
            List<PlanIndexRawItem> raws = _db.FreeSql.Select<PlanIndexRawItem>().Where(r => r.JobId == jobId).ToList();
            if (raws.Count == 0)
            {
                return result;
            }
            // 1. 按测试项目名分组，选出模板文本
            Dictionary<string, List<PlanIndexRawItem>> byName = [];
            foreach (PlanIndexRawItem raw in raws.OrderBy(r => r.EntryId).ThenBy(r => r.OrderNo))
            {
                string key = NameKey(raw.TestItemName);
                if (key.Length == 0)
                {
                    continue;
                }
                if (!byName.TryGetValue(key, out List<PlanIndexRawItem> list))
                {
                    list = [];
                    byName[key] = list;
                }
                list.Add(raw);
            }

            Dictionary<string, PlanItemTemplate> templates = _db.FreeSql.Select<PlanItemTemplate>().ToList()
                .GroupBy(t => NameKey(t.TestItemName))
                .ToDictionary(g => g.Key, g => g.First());
            Dictionary<string, PlanItemTemplate> canonical = [];

            foreach (KeyValuePair<string, List<PlanIndexRawItem>> pair in byName)
            {
                List<PlanIndexRawItem> group = pair.Value;
                string displayName = MostCommon(group.Select(r => r.TestItemName?.Trim()).Where(n => !string.IsNullOrEmpty(n)));
                string category = MostCommon(group.Select(r => r.Category?.Trim()).Where(c => !string.IsNullOrEmpty(c)));
                string sampling = MostCommonNormalized(group.Select(r => r.SamplingPlan));
                string condition = MostCommonNormalized(group.Select(r => r.TestCondition));
                string criterion = MostCommonNormalized(group.Select(r => r.PassCriterion));
                string remark = MostCommonNormalized(group.Select(r => r.Remark));
                int usage = group.Select(r => r.EntryId).Distinct().Count();

                if (!templates.TryGetValue(pair.Key, out PlanItemTemplate template))
                {
                    template = new PlanItemTemplate
                    {
                        TestItemName = displayName ?? pair.Key,
                        Category = category,
                        SamplingPlan = sampling,
                        TestCondition = condition,
                        PassCriterion = criterion,
                        Remark = remark,
                        Period = InferPeriodHours(displayName, condition),
                        UsageCount = usage,
                        CreatedBy = user,
                        CreatedAt = DateTime.Now,
                        UpdatedBy = user,
                        UpdatedAt = DateTime.Now
                    };
                    template.Id = _db.FreeSql.Insert(template).ExecuteIdentity();
                    templates[pair.Key] = template;
                    result.TemplateCount++;
                }
                else
                {
                    template.UsageCount = usage;
                    template.UpdatedBy = user;
                    template.UpdatedAt = DateTime.Now;
                    // 用户手工维护过的模板不覆盖文本，只更新统计
                    if (!template.IsManual)
                    {
                        string period = InferPeriodHours(displayName, condition);
                        bool changed = template.Category != category || template.SamplingPlan != sampling
                            || template.TestCondition != condition || template.PassCriterion != criterion
                            || template.Remark != remark || template.Period != period;
                        if (changed)
                        {
                            template.Category = category;
                            template.SamplingPlan = sampling;
                            template.TestCondition = condition;
                            template.PassCriterion = criterion;
                            template.Remark = remark;
                            template.Period = period;
                        }
                    }
                    _db.FreeSql.Update<PlanItemTemplate>().SetSource(template).Where(t => t.Id == template.Id).ExecuteAffrows();
                }
                canonical[pair.Key] = template;
            }
            result.TemplateCount = canonical.Count;

            // 2. 登记测试项目字典
            HashSet<string> knownItems = _db.FreeSql.Select<TestItemCatalog>().ToList()
                .Select(t => t.Name?.Trim())
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .ToHashSet(StringComparer.CurrentCultureIgnoreCase);
            foreach (PlanItemTemplate template in canonical.Values)
            {
                if (knownItems.Add(template.TestItemName))
                {
                    _db.FreeSql.Insert(new TestItemCatalog
                    {
                        Name = template.TestItemName,
                        Period = template.Period,
                        Remark = "由计划索引自动登记"
                    }).ExecuteAffrows();
                    result.NewTestItems++;
                }
            }

            // 3. 按「机种 + 阶段」生成计划与明细（差异只存不同部分）
            List<PlanIndexEntry> entries = _db.FreeSql.Select<PlanIndexEntry>().Where(e => e.JobId == jobId).ToList();
            Dictionary<long, PlanIndexEntry> entryById = entries.ToDictionary(e => e.Id);
            foreach (IGrouping<string, PlanIndexRawItem> planGroup in raws
                .Where(r => !string.IsNullOrWhiteSpace(r.ModelName))
                .GroupBy(r => $"{r.ModelName}|{PlanStage.Normalize(r.Stage)}"))
            {
                List<PlanIndexRawItem> groupItems = planGroup.ToList();
                string modelName = groupItems[0].ModelName;
                string stage = PlanStage.Normalize(groupItems[0].Stage);
                // 以最新处理的那份报告作为该机种计划的依据
                long latestEntryId = groupItems.Max(r => r.EntryId);
                List<PlanIndexRawItem> latestItems = groupItems
                    .Where(r => r.EntryId == latestEntryId)
                    .OrderBy(r => r.OrderNo)
                    .ToList();

                TestPlan plan = _db.FreeSql.Select<TestPlan>()
                    .Where(p => p.ModelName == modelName && p.Stage == stage).First();
                if (plan == null)
                {
                    plan = new TestPlan
                    {
                        ModelName = modelName,
                        Stage = stage,
                        Source = "Index",
                        CreatedBy = user,
                        CreatedAt = DateTime.Now
                    };
                    plan.Id = _db.FreeSql.Insert(plan).ExecuteIdentity();
                }
                plan.Remark = entryById.TryGetValue(latestEntryId, out PlanIndexEntry latestEntry) ? latestEntry.Note : plan.Remark;
                plan.Source = "Index";
                plan.UpdatedBy = user;
                plan.UpdatedAt = DateTime.Now;
                _db.FreeSql.Update<TestPlan>().SetSource(plan).Where(p => p.Id == plan.Id).ExecuteAffrows();
                result.PlanCount++;

                List<TestPlanItem> existingItems = _db.FreeSql.Select<TestPlanItem>().Where(i => i.PlanId == plan.Id).ToList();
                Dictionary<string, TestPlanItem> existingByName = existingItems
                    .GroupBy(i => NameKey(i.TestItemName))
                    .ToDictionary(g => g.Key, g => g.First());
                HashSet<long> keepIds = [];
                int order = 0;
                foreach (PlanIndexRawItem raw in latestItems)
                {
                    order++;
                    string key = NameKey(raw.TestItemName);
                    if (key.Length == 0 || !canonical.TryGetValue(key, out PlanItemTemplate template))
                    {
                        continue;
                    }
                    existingByName.TryGetValue(key, out TestPlanItem item);
                    bool isNew = item == null;
                    if (isNew)
                    {
                        item = new TestPlanItem { PlanId = plan.Id, TestItemName = raw.TestItemName.Trim() };
                    }
                    // 已人工确认的差异不再被索引覆盖
                    if (isNew || !item.Confirmed)
                    {
                        item.TestItemName = raw.TestItemName.Trim();
                        item.Category = string.IsNullOrWhiteSpace(raw.Category) ? template.Category : raw.Category.Trim();
                        item.TemplateId = template.Id;
                        item.SamplingPlan = Override(raw.SamplingPlan, template.SamplingPlan);
                        item.TestCondition = Override(raw.TestCondition, template.TestCondition);
                        item.PassCriterion = Override(raw.PassCriterion, template.PassCriterion);
                        item.Remark = Override(raw.Remark, template.Remark);
                        item.Period = null; // 周期不在报告里，统一取模板
                    }
                    else if (item.TemplateId != template.Id)
                    {
                        item.TemplateId = template.Id;
                    }
                    item.OrderNo = order;
                    item.FromIndex = true;
                    item.SourceVariants = BuildVariants(groupItems, key, raw, template);
                    item.RefreshOverriddenFields();
                    item.UpdatedBy = user;
                    item.UpdatedAt = DateTime.Now;
                    if (item.HasOverride)
                    {
                        result.DiffCount++;
                    }
                    if (isNew)
                    {
                        item.Id = _db.FreeSql.Insert(item).ExecuteIdentity();
                        existingByName[key] = item;
                    }
                    else
                    {
                        _db.FreeSql.Update<TestPlanItem>().SetSource(item).Where(i => i.Id == item.Id).ExecuteAffrows();
                    }
                    keepIds.Add(item.Id);
                    result.ItemCount++;
                }
                // 清理本机种计划里"来自索引但最新报告已不再包含"的明细；人工新增的保留
                foreach (TestPlanItem item in existingItems.Where(i => i.FromIndex && !keepIds.Contains(i.Id)))
                {
                    _db.FreeSql.Delete<TestPlanItem>().Where(i => i.Id == item.Id).ExecuteAffrows();
                }
            }
            _logger.Info($"计划索引归并完成：模板 {result.TemplateCount} 个，计划 {result.PlanCount} 个，明细 {result.ItemCount} 条，差异 {result.DiffCount} 处");
            return result;
        }

        /// <summary>
        /// 差异取值：与模板（归一化后）一致则返回 null（沿用模板），否则返回原文
        /// </summary>
        private static string Override(string raw, string canonical)
            => OrtPlanParser.Normalize(raw) == OrtPlanParser.Normalize(canonical) ? null : raw;

        /// <summary>
        /// 收集同一机种同一测试项目在其他报告里的不同写法（供人工确认时参考）
        /// </summary>
        private static string BuildVariants(List<PlanIndexRawItem> groupItems, string nameKey,
            PlanIndexRawItem chosen, PlanItemTemplate template)
        {
            List<PlanIndexRawItem> siblings = groupItems.Where(r => NameKey(r.TestItemName) == nameKey).ToList();
            List<string> lines = [];
            Collect(siblings.Select(r => r.SamplingPlan), chosen.SamplingPlan, template.SamplingPlan, "抽样计划");
            Collect(siblings.Select(r => r.TestCondition), chosen.TestCondition, template.TestCondition, "测试条件");
            Collect(siblings.Select(r => r.PassCriterion), chosen.PassCriterion, template.PassCriterion, "通过判定");
            Collect(siblings.Select(r => r.Remark), chosen.Remark, template.Remark, "备注");
            return lines.Count == 0 ? null : string.Join("\n", lines);

            void Collect(IEnumerable<string> values, string chosenText, string templateText, string label)
            {
                string chosenKey = OrtPlanParser.Normalize(chosenText);
                string templateKey = OrtPlanParser.Normalize(templateText);
                HashSet<string> seen = [];
                foreach (string value in values)
                {
                    string key = OrtPlanParser.Normalize(value);
                    if (key.Length == 0 || key == chosenKey || key == templateKey || !seen.Add(key))
                    {
                        continue;
                    }
                    lines.Add($"【{label}】{Shorten(value)}");
                }
            }
        }

        private static string Shorten(string text)
        {
            string single = Regex.Replace(text ?? "", @"\s+", " ").Trim();
            return single.Length <= 160 ? single : single.Substring(0, 160) + "…";
        }

        /// <summary>测试项目名归一化键（大小写/空白无关）</summary>
        public static string NameKey(string name)
            => OrtPlanParser.Normalize(name).Replace("\n", "").ToUpperInvariant();

        /// <summary>出现次数最多的写法（并列时取先出现的）</summary>
        private static string MostCommon(IEnumerable<string> values)
        {
            Dictionary<string, (int Count, int First, string Text)> stat = [];
            int index = 0;
            foreach (string value in values)
            {
                index++;
                if (string.IsNullOrWhiteSpace(value))
                {
                    continue;
                }
                string key = value.Trim();
                if (stat.TryGetValue(key, out (int Count, int First, string Text) item))
                {
                    stat[key] = (item.Count + 1, item.First, item.Text);
                }
                else
                {
                    stat[key] = (1, index, key);
                }
            }
            return stat.Count == 0 ? null
                : stat.OrderByDescending(kv => kv.Value.Count).ThenBy(kv => kv.Value.First).First().Value.Text;
        }

        /// <summary>出现次数最多的文本（按归一化比较，返回原文）</summary>
        private static string MostCommonNormalized(IEnumerable<string> values)
        {
            Dictionary<string, (int Count, int First, string Text)> stat = [];
            int index = 0;
            foreach (string value in values)
            {
                index++;
                string key = OrtPlanParser.Normalize(value);
                if (key.Length == 0)
                {
                    continue;
                }
                if (stat.TryGetValue(key, out (int Count, int First, string Text) item))
                {
                    stat[key] = (item.Count + 1, item.First, item.Text);
                }
                else
                {
                    stat[key] = (1, index, value);
                }
            }
            return stat.Count == 0 ? null
                : stat.OrderByDescending(kv => kv.Value.Count).ThenBy(kv => kv.Value.First).First().Value.Text;
        }

        /// <summary>
        /// 推断试验周期（小时）：优先取条件文本里的 "168 Hrs" 之类，其次按测试项目名给经验值。
        /// 结果只是登记到测试项目字典的初值，用户可改。
        /// </summary>
        public static string InferPeriodHours(string testItemName, string condition)
        {
            Match match = Regex.Match(condition ?? "", @"(\d{1,4})\s*(?:Hrs|Hr|Hours|Hour)\b", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return match.Groups[1].Value;
            }
            string name = (testItemName ?? "").ToLowerInvariant();
            if (name.Contains("burn"))
            {
                return "168";
            }
            return "24";
        }

        /// <summary>切回 UI 线程通知</summary>
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
                _logger.Warn($"计划索引进度通知失败: {ex.Message}");
            }
        }
    }
}
