using NLog;
using ORT一键报告.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace ORT一键报告.Services
{
    /// <summary>
    /// 一次「从平台取看板」的结果（界面提示用）
    /// </summary>
    public class OrtBoardFetchResult
    {
        /// <summary>是否取到数据</summary>
        public bool Success { get; set; }

        /// <summary>失败原因 / 附加说明</summary>
        public string Message { get; set; }

        /// <summary>解析出来的看板行</summary>
        public List<OrtBoardRow> Rows { get; set; } = [];

        /// <summary>实际取数用的控件名</summary>
        public string Widget { get; set; }

        /// <summary>config 返回的控件清单</summary>
        public List<string> Widgets { get; set; } = [];

        /// <summary>原始取数响应字节数（排查用）</summary>
        public int RawLength { get; set; }

        /// <summary>取数条件摘要（"20261101~20261126 类型=全部"）</summary>
        public string QueryText { get; set; }
    }

    /// <summary>
    /// ORT 看板取数：把「公司平台登录 → 帆软看板 → 取数 JSON」串起来，
    /// 还原成和 Excel 导入完全一样的 <see cref="OrtBoardRow"/>，再交给 <see cref="PreIssueService"/> 入库。
    /// 这样「从平台获取」和「导入 Excel」走同一条入库/比对/变动记录逻辑，行为一致。
    /// </summary>
    public class OrtBoardService
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;
        private readonly AuthService _auth;
        private readonly OrtBoardSettingsService _settings;
        private readonly PortalAccountService _portal;
        private readonly FineReportClient _client;
        private readonly PreIssueService _preIssue;

        public OrtBoardService(DatabaseService db, AuthService auth, OrtBoardSettingsService settings,
            PortalAccountService portal, FineReportClient client, PreIssueService preIssue)
        {
            _db = db;
            _auth = auth;
            _settings = settings;
            _portal = portal;
            _client = client;
            _preIssue = preIssue;
        }

        /* ###############################  取数  ################################ */

        /// <summary>
        /// 用当前登录的平台账号（或最近登记的账号）登录并取一次看板。
        /// 失败不抛异常，结果里给 Message。
        /// </summary>
        public async Task<OrtBoardFetchResult> FetchAsync(CancellationToken ct = default)
        {
            OrtBoardFetchResult result = new();
            OrtBoardSettings current = _settings.Get();
            result.QueryText = DescribeQuery(current.Query);
            try
            {
                if (!current.IsQueryReady(out string queryError))
                {
                    result.Message = queryError;
                    return result;
                }
                if (!current.IsTargetReady(out string targetError))
                {
                    result.Message = targetError;
                    return result;
                }

                PortalAccount account = ResolveFetchAccount(out string accountError);
                if (account == null)
                {
                    result.Message = accountError;
                    return result;
                }

                FineReportSession session = await GetSessionAsync(account, ct).ConfigureAwait(false);
                FineReportTable table = await _client.FetchAsync(session, current.Endpoints, current.Query, ct)
                    .ConfigureAwait(false);

                result.Widget = table.Widget;
                result.Widgets = table.Widgets;
                result.RawLength = table.RawLength;
                result.Rows = ToBoardRows(table.Rows, _auth.CurrentOperatorName, out string parseError);
                if (result.Rows.Count == 0)
                {
                    // 带上原始字节数：如果日志里响应很大却 0 行，说明是解析问题而不是没取到数据
                    result.Message = (parseError
                        ?? "取到了报表数据，但里面没有认得出的看板行（核对高级选项里的 TargetEntryId / WidgetHint 是不是这张看板）")
                        + $"（原始响应 {table.RawLength} 字节）";
                    return result;
                }
                result.Success = true;
                result.Message = $"控件 {table.Widget}，{result.Rows.Count} 行";
                _logger.Info($"从平台取看板成功：{result.Message}（{result.QueryText}）");
                return result;
            }
            catch (PortalAuthException ex)
            {
                _logger.Warn($"从平台取看板失败（{ex.Status}）：{ex.Message}");
                result.Message = ex.Message;
                return result;
            }
            catch (OperationCanceledException)
            {
                result.Message = "已取消取数";
                return result;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "从平台取看板失败");
                result.Message = "取数失败：" + ex.Message;
                return result;
            }
        }

        /// <summary>
        /// 取数用哪个账号：当前登录的平台账号 → 唯一的已登记账号 → 报错（提示先去登记/登录）。
        /// </summary>
        private PortalAccount ResolveFetchAccount(out string error)
        {
            error = null;
            long userId = _auth.CurrentUser?.Id ?? 0;
            if (userId > 0)
            {
                PortalAccount mine = _portal.FindByUser(userId);
                if (mine != null && mine.IsActive)
                {
                    return mine;
                }
            }
            List<PortalAccount> all = _portal.ListAccounts().Where(a => a.IsActive).ToList();
            if (all.Count == 1)
            {
                return all[0];
            }
            error = all.Count == 0
                ? "还没有登记过公司平台账号：请先用公司平台账号登录，或在「设置 → 登录与身份 → 部门账号」里登记"
                : "当前登录不是公司平台账号，无法确定用哪个账号取数：请用公司平台账号登录后再试";
            return null;
        }

        /// <summary>拿到可用会话：先用缓存，过期就用保存的密码重新登录</summary>
        private async Task<FineReportSession> GetSessionAsync(PortalAccount account, CancellationToken ct)
        {
            FineReportSession cached = _portal.GetCachedSession(account.Account);
            if (cached != null && cached.IsTokenUsable)
            {
                return cached;
            }
            string password = PortalAccountService.GetStoredPassword(account);
            if (string.IsNullOrWhiteSpace(password))
            {
                throw new PortalAuthException(PortalLoginStatus.EmptyCredentials,
                    $"账号 {account.Account} 没有保存可用的密码，请在登录窗口用该账号登录一次");
            }
            OrtBoardEndpoints ep = _settings.Get().Endpoints;
            string deviceUuid = _settings.EnsureDeviceUuid(ep);
            FineReportSession session = await _client.LoginAsync(account.Account, password, deviceUuid, ep, ct)
                .ConfigureAwait(false);
            _logger.Info($"已为取数重新登录平台：{account.Account}");
            return session;
        }

        /// <summary>
        /// 预览（不写库）：给用户确认「哪个月、多少行、会覆盖多少」。
        /// </summary>
        public PreIssueImportPreview Preview(IList<OrtBoardRow> rows, string sourceName)
        {
            PreIssueImportPreview preview = new()
            {
                Source = PreIssueCodes.SourceBoard,
                File = sourceName
            };
            if (rows == null || rows.Count == 0)
            {
                preview.Success = false;
                preview.Message = "没有读到看板数据";
                return preview;
            }
            List<string> months = rows.Select(r => r.Month).Where(m => !string.IsNullOrWhiteSpace(m))
                .Distinct().OrderBy(m => m, StringComparer.Ordinal).ToList();
            preview.Month = months.FirstOrDefault();
            preview.Items = rows.Count;
            preview.Models = rows.Select(r => r.Model).Where(m => !string.IsNullOrWhiteSpace(m))
                .Distinct(StringComparer.OrdinalIgnoreCase).Count();
            preview.Existing = months.Sum(m => (int)_db.FreeSql.Select<OrtBoardRow>().Where(a => a.Month == m).Count());
            int stocked = rows.Count(r => PreIssueCodes.IsYes(r.Stocked87));
            string monthText = months.Count > 1
                ? string.Join("、", months.Select(PreIssueService.MonthText))
                : PreIssueService.MonthText(preview.Month);
            preview.Sample = $"  {monthText}：{rows.Count} 行，其中 SIM87入庫否 = Y 的 {stocked} 行"
                + "\n" + string.Join("\n", rows.Take(2).Select(r =>
                    $"  {r.ReportDate}  {r.Model}  {r.JobNo}  SIM87={r.Stocked87}"));
            return preview;
        }

        /// <summary>写入数据库（复用 Excel 导入同一套比对/变动记录逻辑）</summary>
        public PreIssueImportResult Import(IList<OrtBoardRow> rows, string sourceName)
            => _preIssue.ImportBoardRows(rows, sourceName);

        /// <summary>
        /// 动态读取控件清单（高级选项里 WidgetHint 的动态获取）：
        /// 用当前账号登录 → 打开 TargetEntryId 这张看板 → 读 config。
        /// </summary>
        public async Task<(bool Ok, string Message, List<string> Widgets)> ProbeWidgetsAsync(
            CancellationToken ct = default)
        {
            try
            {
                OrtBoardSettings current = _settings.Get();
                if (!current.IsTargetReady(out string targetError))
                {
                    return (false, targetError, []);
                }
                PortalAccount account = ResolveFetchAccount(out string accountError);
                if (account == null)
                {
                    return (false, accountError, []);
                }
                FineReportSession session = await GetSessionAsync(account, ct).ConfigureAwait(false);
                List<string> widgets = await _client.ProbeWidgetsAsync(session, current.Endpoints, ct)
                    .ConfigureAwait(false);
                if (widgets.Count == 0)
                {
                    return (false, "这张看板没有返回控件清单，请核对 TargetEntryId", widgets);
                }
                return (true, null, widgets);
            }
            catch (PortalAuthException ex)
            {
                return (false, ex.Message, []);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "读取控件列表失败");
                return (false, "读取控件列表失败：" + ex.Message, []);
            }
        }

        /* ###############################  解析  ################################ */

        /// <summary>
        /// 帆软单元格矩阵 → 看板行。表头行按列名认（日期/類別/機種/生產線別/生產日期/工令…），
        /// 只有同时认出「機種」且列名命中 ≥4 个才算表头，避免把标题行当表头。
        /// </summary>
        public static List<OrtBoardRow> ToBoardRows(List<List<string>> grid, string operatorName, out string error)
        {
            error = null;
            List<OrtBoardRow> rows = [];
            if (grid == null || grid.Count == 0)
            {
                error = "报表返回的内容是空的";
                return rows;
            }

            int headerRow = FindHeaderRow(grid, out Dictionary<string, int> map);
            if (headerRow < 0)
            {
                error = "没认出看板表头（應有 日期/類別/機種/生產線別/生產日期/工令…），"
                    + "说明取到的可能不是这张看板，请核对 TargetEntryId / WidgetHint";
                return rows;
            }

            DateTime now = DateTime.Now;
            for (int index = headerRow + 1; index < grid.Count; index++)
            {
                List<string> line = grid[index];
                string reportDate = Cell(line, map, "日期");
                string model = Cell(line, map, "機種", "机种");
                string jobNo = Cell(line, map, "工令");
                if (string.IsNullOrWhiteSpace(model) && string.IsNullOrWhiteSpace(jobNo))
                {
                    continue;
                }
                string month = PreIssueService.MonthFromYmd(reportDate)
                    ?? PreIssueService.MonthFromYmd(Cell(line, map, "生產日期", "生产日期"));
                if (month == null)
                {
                    continue;
                }
                rows.Add(new OrtBoardRow
                {
                    Month = month,
                    ReportDate = reportDate,
                    Category = Cell(line, map, "類別", "类别"),
                    Model = model,
                    Line = Cell(line, map, "生產線別", "生产线别"),
                    ProduceDate = Cell(line, map, "生產日期", "生产日期"),
                    JobNo = jobNo,
                    Qty = Cell(line, map, "生產數量", "生产数量"),
                    Hours = Cell(line, map, "生產時數", "生产时数"),
                    Reported65 = Cell(line, map, "65報工否", "65报工否"),
                    Stocked87 = Cell(line, map, "SIM87入庫否", "SIM87入库否"),
                    ReceiveDate = Cell(line, map, "收料日期"),
                    PartNo = Cell(line, map, "料號", "料号"),
                    MaterialStock = Cell(line, map, "材料剩餘庫存", "材料剩余库存"),
                    TestResult = Cell(line, map, "測試結果", "测试结果"),
                    Forbidden94 = Cell(line, map, "SIM94禁止出貨", "SIM94禁止出货"),
                    ImportedAt = now,
                    ImportedBy = operatorName
                });
            }
            if (rows.Count == 0)
            {
                error = "看板表头认出来了，但下面没有数据行（日期区间可能选错了）";
            }
            return rows;
        }

        /// <summary>看板列名（用来认表头）</summary>
        private static readonly string[] KnownHeaders =
        [
            "日期", "類別", "类别", "機種", "机种", "生產線別", "生产线别", "生產日期", "生产日期",
            "工令", "生產數量", "生产数量", "生產時數", "生产时数", "65報工否", "65报工否",
            "SIM87入庫否", "SIM87入库否", "收料日期", "料號", "料号", "材料剩餘庫存", "材料剩余库存",
            "測試結果", "测试结果", "SIM94禁止出貨", "SIM94禁止出货"
        ];

        /// <summary>在前几行里找表头，返回行号（-1 = 没找到）与列名→列号映射</summary>
        private static int FindHeaderRow(List<List<string>> grid, out Dictionary<string, int> map)
        {
            map = [];
            int limit = Math.Min(grid.Count, 8);
            for (int row = 0; row < limit; row++)
            {
                Dictionary<string, int> candidate = [];
                bool hasModel = false;
                int hits = 0;
                List<string> line = grid[row];
                for (int col = 0; col < line.Count; col++)
                {
                    string header = (line[col] ?? "").Trim();
                    if (header.Length == 0)
                    {
                        continue;
                    }
                    if (Array.IndexOf(KnownHeaders, header) >= 0)
                    {
                        hits++;
                        if (!candidate.ContainsKey(header))
                        {
                            candidate[header] = col;
                        }
                    }
                    if (header == "機種" || header == "机种")
                    {
                        hasModel = true;
                    }
                }
                if (hasModel && hits >= 4)
                {
                    map = candidate;
                    return row;
                }
            }
            return -1;
        }

        /// <summary>按列名取单元格文本（多个别名依次尝试，空则跳过）</summary>
        private static string Cell(List<string> line, Dictionary<string, int> map, params string[] headers)
        {
            foreach (string header in headers)
            {
                if (map.TryGetValue(header, out int col) && col < line.Count)
                {
                    string text = line[col];
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        return text.Trim();
                    }
                }
            }
            return null;
        }

        /// <summary>取数条件摘要（日志与提示用）</summary>
        public static string DescribeQuery(OrtBoardQuery query)
        {
            if (query == null)
            {
                return "(未设置)";
            }
            string moType = string.IsNullOrWhiteSpace(query.MoType) ? "全部" : query.MoType;
            return $"{query.DateFrom}~{query.DateTo} 类型={moType}";
        }
    }
}
