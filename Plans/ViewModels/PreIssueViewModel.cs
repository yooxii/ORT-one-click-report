using CommunityToolkit.Mvvm.ComponentModel;
using NLog;
using ORT一键报告.Models;
using ORT一键报告.Services;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;

namespace ORT一键报告.Plans.ViewModels
{
    /// <summary>
    /// 预领（预领标签页）的视图模型：三个内层页签共用一个月份，
    /// 1) 月排程：平铺矩阵的周视图 / 整月视图，导入更新并记录变动；
    /// 2) ORT看板：原始导出数据（表头在第 2 行、只取 A~O），每天导入更新，比对 SIM87入庫否；
    /// 3) 准备领用：先填待开单机种，导入开单数据后补全 10 列，再次导入同步 VCH_STATUS / CLOSE_CODE。
    /// </summary>
    public class PreIssueViewModel : ObservableObject
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly PreIssueService _service;
        private readonly IPermissionService _permission;
        private readonly IPathService _pathService;
        private readonly OrtBoardService _board;

        public PreIssueViewModel(PreIssueService service, IPermissionService permission, IPathService pathService,
            OrtBoardService board)
        {
            _service = service;
            _permission = permission;
            _pathService = pathService;
            _board = board;
            BoardView = CollectionViewSource.GetDefaultView(BoardRows);
            BoardView.Filter = FilterBoard;
            ReqView = CollectionViewSource.GetDefaultView(ReqRows);
            ReqView.Filter = FilterRequisition;
            BorrowView = CollectionViewSource.GetDefaultView(BorrowRows);
            BorrowView.Filter = FilterBorrow;
            LoadMonths();
        }

        /* ###############################  月份  ################################ */

        public ObservableCollection<MonthOption> Months { get; } = [];

        private MonthOption _selectedMonth;
        /// <summary>当前查看的月份（三个页签共用；旧月份照样可以翻看、也可以再次导入更新）</summary>
        public MonthOption SelectedMonth
        {
            get => _selectedMonth;
            set
            {
                if (SetProperty(ref _selectedMonth, value))
                {
                    OnPropertyChanged(nameof(MonthTitle));
                    ReloadAll();
                }
            }
        }

        /// <summary>当前月份文本（2026/11）</summary>
        public string MonthTitle => SelectedMonth?.Text ?? PreIssueService.MonthText(PreIssueService.CurrentMonth);

        private string Month => SelectedMonth?.Month ?? PreIssueService.CurrentMonth;

        /// <summary>重新读取月份列表（导入后可能新增月份）</summary>
        public void LoadMonths()
        {
            string keep = SelectedMonth?.Month ?? PreIssueService.CurrentMonth;
            Months.Clear();
            foreach (string month in _service.GetMonths())
            {
                Months.Add(new MonthOption { Month = month, Text = PreIssueService.MonthText(month) });
            }
            SelectedMonth = Months.FirstOrDefault(m => m.Month == keep) ?? Months.FirstOrDefault();
        }

        /* ###############################  月排程  ################################ */

        /// <summary>月排程的行（一个「線別 + 機種」一行，格子是当天数量）</summary>
        public ObservableCollection<PreScheduleRow> ScheduleRows { get; } = [];

        /// <summary>月排程当前显示的日列（周视图 7 列 / 整月视图当月天数）</summary>
        public ObservableCollection<ScheduleDayColumn> ScheduleColumns { get; } = [];

        /// <summary>周列表（月排程按周翻页，避免左右横移）</summary>
        public ObservableCollection<WeekOption> Weeks { get; } = [];

        private WeekOption _selectedWeek;
        public WeekOption SelectedWeek
        {
            get => _selectedWeek;
            set
            {
                if (SetProperty(ref _selectedWeek, value) && !WholeMonth)
                {
                    BuildScheduleColumns();
                }
            }
        }

        private bool _wholeMonth;
        /// <summary>整月视图（平铺 1 日…31 日）；关闭时是周视图（一次 7 天，左右翻周）</summary>
        public bool WholeMonth
        {
            get => _wholeMonth;
            set
            {
                if (SetProperty(ref _wholeMonth, value))
                {
                    OnPropertyChanged(nameof(WeekView));
                    BuildScheduleColumns();
                }
            }
        }

        /// <summary>周视图（与 <see cref="WholeMonth"/> 互斥，方便绑定单选按钮）</summary>
        public bool WeekView
        {
            get => !WholeMonth;
            set => WholeMonth = !value;
        }

        private string _scheduleSearch;
        /// <summary>排程搜索（按線別/機種过滤）</summary>
        public string ScheduleSearch
        {
            get => _scheduleSearch;
            set
            {
                if (SetProperty(ref _scheduleSearch, value))
                {
                    BuildScheduleRows();
                }
            }
        }

        private string _scheduleSummary;
        public string ScheduleSummary { get => _scheduleSummary; private set => SetProperty(ref _scheduleSummary, value); }

        private string _scheduleHint;
        /// <summary>排程的导入提示（最后导入时间 / 是否该更新了）</summary>
        public string ScheduleHint { get => _scheduleHint; private set => SetProperty(ref _scheduleHint, value); }

        /// <summary>列头变化时通知界面重建表格列（日列数量随周/整月变化）</summary>
        public event Action ScheduleColumnsChanged;

        private List<PreScheduleItem> _scheduleItems = [];

        private void LoadSchedule()
        {
            try
            {
                _scheduleItems = _service.LoadSchedule(Month);
            }
            catch (Exception ex)
            {
                _logger.Warn($"读取月排程失败: {ex.Message}");
                _scheduleItems = [];
            }
            BuildWeeks();
            BuildScheduleColumns();
            DateTime? last = _service.LastImportAt(PreIssueCodes.SourceSchedule, Month);
            ScheduleHint = BuildScheduleHint(last);
        }

        private string BuildScheduleHint(DateTime? last)
        {
            if (_scheduleItems.Count == 0)
            {
                return "本月还没有排程数据：点「导入更新」选择排程表（年月 / 線別 / 1日…31日，单元格写「機種/數量」）";
            }
            string when = last.HasValue ? last.Value.ToString("yyyy/M/d HH:mm", CultureInfo.InvariantCulture) : "未知";
            bool overdue = !last.HasValue || (DateTime.Now - last.Value).TotalDays >= 8;
            return overdue
                ? $"最后导入：{when} —— 排程随时会调整，建议重新导入最新排程"
                : $"最后导入：{when}";
        }

        private void BuildWeeks()
        {
            DateTime? start = PreIssueService.MonthStart(Month);
            int keepIndex = SelectedWeek?.Index ?? -1;
            Weeks.Clear();
            if (start == null)
            {
                SelectedWeek = null;
                return;
            }
            DateTime first = start.Value;
            int daysInMonth = DateTime.DaysInMonth(first.Year, first.Month);
            int offset = ((int)first.DayOfWeek + 6) % 7; // 周一为一周起点
            DateTime cursor = first.AddDays(-offset);
            int index = 0;
            while (cursor < first.AddDays(daysInMonth))
            {
                DateTime weekStart = cursor;
                Weeks.Add(new WeekOption
                {
                    Index = index++,
                    Start = weekStart,
                    Text = $"{weekStart:M/d}～{weekStart.AddDays(6):M/d}"
                });
                cursor = cursor.AddDays(7);
            }
            SelectedWeek = Weeks.FirstOrDefault(w => w.Index == keepIndex) ?? PickCurrentWeek() ?? Weeks.FirstOrDefault();
        }

        /// <summary>当前月份默认选中「今天所在的那一周」，历史月份默认第一周</summary>
        private WeekOption PickCurrentWeek()
        {
            DateTime today = DateTime.Today;
            if (SelectedMonth?.Month != PreIssueService.CurrentMonth)
            {
                return null;
            }
            return Weeks.FirstOrDefault(w => today >= w.Start && today <= w.Start.AddDays(6));
        }

        private void BuildScheduleColumns()
        {
            ScheduleColumns.Clear();
            DateTime? start = PreIssueService.MonthStart(Month);
            if (start == null)
            {
                ScheduleColumnsChanged?.Invoke();
                return;
            }
            DateTime first = start.Value;
            int daysInMonth = DateTime.DaysInMonth(first.Year, first.Month);
            if (WholeMonth)
            {
                for (int day = 1; day <= daysInMonth; day++)
                {
                    AddDayColumn(first, day, true);
                }
            }
            else if (SelectedWeek != null)
            {
                for (int i = 0; i < 7; i++)
                {
                    DateTime date = SelectedWeek.Start.AddDays(i);
                    bool inMonth = date.Month == first.Month && date.Year == first.Year;
                    AddDayColumn(first, date.Day, inMonth);
                }
            }
            BuildScheduleRows();
            ScheduleColumnsChanged?.Invoke();
        }

        private void AddDayColumn(DateTime first, int day, bool inMonth)
        {
            DateTime date = new(first.Year, first.Month, Math.Min(day, DateTime.DaysInMonth(first.Year, first.Month)));
            string week = "日一二三四五六"[(int)date.DayOfWeek].ToString();
            ScheduleColumns.Add(new ScheduleDayColumn
            {
                Day = inMonth ? day : 0,
                Header = inMonth ? $"{day}日 {week}" : "-",
                IsWeekend = inMonth && (date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday),
                InMonth = inMonth
            });
        }

        private void BuildScheduleRows()
        {
            ScheduleRows.Clear();
            string keyword = ScheduleSearch?.Trim();
            IEnumerable<IGrouping<string, PreScheduleItem>> groups = _scheduleItems
                .Where(i => !string.IsNullOrWhiteSpace(i.Model))
                .GroupBy(i => (i.Line ?? "").Trim() + "\u0001" + i.Model.Trim());

            List<PreScheduleRow> rows = [];
            foreach (IGrouping<string, PreScheduleItem> group in groups)
            {
                List<PreScheduleItem> items = group.ToList();
                PreScheduleItem first = items[0];
                PreScheduleRow row = new()
                {
                    Line = first.Line,
                    Model = first.Model,
                    Days = new string[32],
                    Quantities = new int[32]
                };
                int total = 0;
                foreach (PreScheduleItem item in items)
                {
                    if (item.Day < 1 || item.Day > 31)
                    {
                        continue;
                    }
                    row.Quantities[item.Day] += item.Qty;
                    total += item.Qty;
                }
                for (int day = 1; day <= 31; day++)
                {
                    row.Days[day] = row.Quantities[day] > 0 ? row.Quantities[day].ToString("N0", CultureInfo.InvariantCulture) : "";
                }
                row.Total = total;
                if (keyword != null && keyword.Length > 0
                    && (row.Line ?? "").IndexOf(keyword, StringComparison.OrdinalIgnoreCase) < 0
                    && (row.Model ?? "").IndexOf(keyword, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }
                // 线别 / 机种列头筛选（多选，命中任意一个即显示）
                if (!ScheduleColumnAllows(nameof(PreScheduleRow.Line), row.Line)
                    || !ScheduleColumnAllows(nameof(PreScheduleRow.Model), row.Model))
                {
                    continue;
                }
                rows.Add(row);
            }
            foreach (PreScheduleRow row in rows
                .OrderBy(r => r.Line, StringComparer.OrdinalIgnoreCase)
                .ThenBy(r => r.Model, StringComparer.OrdinalIgnoreCase))
            {
                ScheduleRows.Add(row);
            }
            ScheduleSummary = _scheduleItems.Count == 0
                ? "没有排程数据"
                : $"{ScheduleRows.Count} 个機種 · {_scheduleItems.Count} 个排班格 · 合计 {_scheduleItems.Sum(i => (long)i.Qty):N0} pcs";
        }

        /* ------------------ 月排程：线别 / 机种列头筛选 ------------------ */

        /// <summary>月排程的列筛选（属性名 → 勾选的值；目前只有線別 / 機種两列）</summary>
        private readonly Dictionary<string, HashSet<string>> _scheduleColumnFilters = new(StringComparer.Ordinal);

        private bool ScheduleColumnAllows(string property, string value) =>
            !_scheduleColumnFilters.TryGetValue(property, out HashSet<string> set)
            || set.Contains((value ?? "").Trim());

        /// <summary>这一列有没有筛选</summary>
        public bool HasScheduleColumnFilter(string property) =>
            !string.IsNullOrEmpty(property) && _scheduleColumnFilters.ContainsKey(property);

        /// <summary>某一列当前勾选的值（没有筛选时返回 null = 全部显示）</summary>
        public HashSet<string> GetScheduleColumnSelection(string property) =>
            !string.IsNullOrEmpty(property) && _scheduleColumnFilters.TryGetValue(property, out HashSet<string> set)
                ? new HashSet<string>(set, StringComparer.Ordinal)
                : null;

        /// <summary>取一列的去重值（列头筛选菜单用）</summary>
        public List<string> GetScheduleColumnValues(string property)
        {
            bool byLine = property == nameof(PreScheduleRow.Line);
            if (!byLine && property != nameof(PreScheduleRow.Model))
            {
                return [];
            }
            return _scheduleItems
                .Select(i => ((byLine ? i.Line : i.Model) ?? "").Trim())
                .Distinct(StringComparer.Ordinal)
                .OrderBy(v => v, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        /// <summary>设置某一列的筛选：没勾 / 全部勾上 都等于「不筛选」</summary>
        public void SetScheduleColumnFilter(string property, IEnumerable<string> values)
        {
            if (string.IsNullOrEmpty(property))
            {
                return;
            }
            HashSet<string> set = new((values ?? []).Select(v => (v ?? "").Trim()), StringComparer.Ordinal);
            List<string> all = GetScheduleColumnValues(property);
            bool everything = all.Count == 0 || (set.Count >= all.Count && all.All(set.Contains));
            if (set.Count == 0 || everything)
            {
                _scheduleColumnFilters.Remove(property);
            }
            else
            {
                _scheduleColumnFilters[property] = set;
            }
            BuildScheduleRows();
        }

        /// <summary>清除某一列筛选；property 为空时清除全部（线别 + 机种）</summary>
        public void ClearScheduleColumnFilter(string property)
        {
            if (string.IsNullOrEmpty(property))
            {
                _scheduleColumnFilters.Clear();
            }
            else
            {
                _scheduleColumnFilters.Remove(property);
            }
            BuildScheduleRows();
        }

        /* ###############################  ORT 看板  ################################ */

        public ObservableCollection<OrtBoardRow> BoardRows { get; } = [];

        public ICollectionView BoardView { get; }

        private OrtBoardRow _selectedBoardRow;
        public OrtBoardRow SelectedBoardRow { get => _selectedBoardRow; set => SetProperty(ref _selectedBoardRow, value); }

        private string _boardSearch;
        public string BoardSearch
        {
            get => _boardSearch;
            set
            {
                if (SetProperty(ref _boardSearch, value))
                {
                    RefreshView(BoardView);
                }
            }
        }

        private bool _boardOnlyStocked;
        /// <summary>只看已入库（SIM87入庫否 = Y）的行</summary>
        public bool BoardOnlyStocked
        {
            get => _boardOnlyStocked;
            set
            {
                if (SetProperty(ref _boardOnlyStocked, value))
                {
                    RefreshView(BoardView);
                }
            }
        }

        private bool _boardOnlyPending;
        /// <summary>只看还没转准备领用的已入库行</summary>
        public bool BoardOnlyPending
        {
            get => _boardOnlyPending;
            set
            {
                if (SetProperty(ref _boardOnlyPending, value))
                {
                    RefreshView(BoardView);
                }
            }
        }

        private string _boardSummary;
        public string BoardSummary { get => _boardSummary; private set => SetProperty(ref _boardSummary, value); }

        private string _boardHint;
        public string BoardHint { get => _boardHint; private set => SetProperty(ref _boardHint, value); }

        private List<OrtBoardRow> _boardAll = [];

        private void LoadBoard()
        {
            try
            {
                _boardAll = _service.LoadBoard(Month);
            }
            catch (Exception ex)
            {
                _logger.Warn($"读取 ORT 看板失败: {ex.Message}");
                _boardAll = [];
            }
            BoardRows.Clear();
            foreach (OrtBoardRow row in _boardAll)
            {
                BoardRows.Add(row);
            }
            RefreshView(BoardView);
            DateTime? last = _service.LastImportAt(PreIssueCodes.SourceBoard, Month);
            string when = last.HasValue ? last.Value.ToString("yyyy/M/d HH:mm", CultureInfo.InvariantCulture) : "未知";
            bool overdue = !last.HasValue || last.Value.Date < DateTime.Today;
            BoardHint = _boardAll.Count == 0
                ? "本月还没有看板数据：点「导入更新」选择看板原始导出（第 2 行是表头，只取前 15 列）"
                : overdue
                    ? $"最后导入：{when} —— 看板每天早上都要更新，建议导入今早的数据"
                    : $"最后导入：{when}";
            UpdateBoardSummary();
        }

        private void UpdateBoardSummary()
        {
            int total = _boardAll.Count;
            int stocked = _boardAll.Count(r => PreIssueCodes.IsYes(r.Stocked87));
            int pending = _boardAll.Count(r => PreIssueCodes.IsYes(r.Stocked87) && !r.HasRequisition);
            int forbidden = _boardAll.Count(r => PreIssueCodes.IsYes(r.Forbidden94));
            int blacklisted = _boardAll.Count(r => r.Blacklisted);
            BoardSummary = total == 0
                ? "没有看板数据"
                : $"{total} 行 · 已入库 {stocked} · 待转领用 {pending} · SIM94禁止出货 {forbidden}"
                    + (blacklisted > 0 ? $" · 黑名单 {blacklisted}" : "");
        }

        private bool FilterBoard(object obj)
        {
            if (obj is not OrtBoardRow row)
            {
                return false;
            }
            if (BoardOnlyStocked && !PreIssueCodes.IsYes(row.Stocked87))
            {
                return false;
            }
            if (BoardOnlyPending && !(PreIssueCodes.IsYes(row.Stocked87) && !row.HasRequisition))
            {
                return false;
            }
            string keyword = BoardSearch?.Trim();
            if (string.IsNullOrEmpty(keyword))
            {
                return true;
            }
            return Contains(row.Model, keyword) || Contains(row.JobNo, keyword) || Contains(row.PartNo, keyword)
                || Contains(row.Line, keyword) || Contains(row.ReportDate, keyword);
        }

        /* ###############################  准备领用  ################################ */

        public ObservableCollection<PreIssueRequisition> ReqRows { get; } = [];

        public ICollectionView ReqView { get; }

        private PreIssueRequisition _selectedReqRow;
        public PreIssueRequisition SelectedReqRow { get => _selectedReqRow; set => SetProperty(ref _selectedReqRow, value); }

        private string _reqSearch;
        public string ReqSearch
        {
            get => _reqSearch;
            set
            {
                if (SetProperty(ref _reqSearch, value))
                {
                    RefreshView(ReqView);
                }
            }
        }

        /// <summary>状态筛选（全部 / 待开单 / 已开单）</summary>
        public ObservableCollection<string> ReqStateOptions { get; } = [];

        private string _reqStateFilter;
        public string ReqStateFilter
        {
            get => _reqStateFilter;
            set
            {
                if (SetProperty(ref _reqStateFilter, value))
                {
                    RefreshView(ReqView);
                }
            }
        }

        private string _reqSummary;
        public string ReqSummary { get => _reqSummary; private set => SetProperty(ref _reqSummary, value); }

        private List<PreIssueRequisition> _reqAll = [];

        private void LoadRequisition()
        {
            if (ReqStateOptions.Count == 0)
            {
                ReqStateOptions.Add("全部");
                ReqStateOptions.Add(PreIssueCodes.PendingState);
                ReqStateOptions.Add(PreIssueCodes.OpenedState);
                _reqStateFilter = "全部";
            }
            try
            {
                _reqAll = _service.LoadRequisition(Month);
            }
            catch (Exception ex)
            {
                _logger.Warn($"读取准备领用失败: {ex.Message}");
                _reqAll = [];
            }
            ReqRows.Clear();
            foreach (PreIssueRequisition row in _reqAll)
            {
                ReqRows.Add(row);
            }
            RefreshView(ReqView);
            UpdateReqSummary();
        }

        private bool FilterRequisition(object obj)
        {
            if (obj is not PreIssueRequisition row)
            {
                return false;
            }
            bool pending = PreIssueCodes.IsPendingState(row.State) || string.IsNullOrWhiteSpace(row.Vchno);
            if (ReqStateFilter == PreIssueCodes.PendingState && !pending)
            {
                return false;
            }
            if (ReqStateFilter == PreIssueCodes.OpenedState && pending)
            {
                return false;
            }
            // 列头筛选：每列勾选的值里命中任意一个就算通过
            foreach (KeyValuePair<string, HashSet<string>> filter in _reqColumnFilters)
            {
                System.Reflection.PropertyInfo info = typeof(PreIssueRequisition).GetProperty(filter.Key);
                if (info == null)
                {
                    continue;
                }
                string value = info.GetValue(row) as string ?? "";
                if (!filter.Value.Contains(value))
                {
                    return false;
                }
            }
            string keyword = ReqSearch?.Trim();
            if (string.IsNullOrEmpty(keyword))
            {
                return true;
            }
            return Contains(row.PartNo, keyword) || Contains(row.Vchno, keyword) || Contains(row.VchStatus, keyword)
                || Contains(row.JobNo, keyword) || Contains(row.CloseCode, keyword) || Contains(row.Pc, keyword);
        }

        /// <summary>变动记录（此月份、指定来源；来源为空表示该月全部来源）</summary>
        public List<PreIssueChangeLog> LoadChangesForView(string source) => _service.LoadChanges(Month, source);

        /// <summary>变动记录（指定月份、指定来源）</summary>
        public List<PreIssueChangeLog> LoadChangesForMonth(string month, string source) =>
            _service.LoadChanges(month, source);

        /// <summary>变动记录（全部月份，最新在前）</summary>
        public List<PreIssueChangeLog> LoadAllChangesForView(int limit = 800) => _service.LoadChanges(null, null, limit);

        /// <summary>黑名单（管理窗口用）</summary>
        public List<OrtBlacklistItem> LoadBlacklist() => _service.LoadBlacklist();

        /// <summary>加入黑名单（黑名单管理窗口用；isPattern = true 按正则模式）</summary>
        public bool AddBlacklistModel(string model, string reason, bool isPattern = false) =>
            _service.AddToBlacklist(model, reason, isPattern, out string error) || Fail(error);

        /// <summary>校验正则写法（黑名单窗口在加入前先查一遍）</summary>
        public static bool ValidateBlacklistPattern(string pattern, out string error) =>
            OrtBlacklistMatcher.TryCompile(pattern, out _, out error);

        /// <summary>移出黑名单</summary>
        public bool RemoveBlacklistModel(string model) =>
            _service.RemoveFromBlacklist(model, out string error) || Fail(error);

        private static bool Fail(string error)
        {
            if (!string.IsNullOrWhiteSpace(error))
            {
                _ = MessageBox.Show(error, LanguageService.Get("PreIssue_BlacklistTitle"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            return false;
        }

        /// <summary>重新读取黑名单相关状态（黑名单变动后刷新看板标记）</summary>
        public void ReloadAfterBlacklistChange() => ReloadAll();

        /// <summary>添加待开单窗口用的候选机种</summary>
        public List<string> PendingModelCandidates() => BoardModelCandidates();

        /* ###############################  公共状态  ################################ */

        private string _statusMessage;
        public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }

        /// <summary>导入/更新需要导入权限</summary>
        public bool CanImport => _permission.Can("plan.import");

        /// <summary>当前是否有数据可以翻阅</summary>
        public bool HasAnyData => _scheduleItems.Count > 0 || _boardAll.Count > 0 || _reqAll.Count > 0 || _borrowAll.Count > 0;

        /// <summary>总体提示（预领页签顶部的一行说明）</summary>
        public string OverallHint
        {
            get
            {
                string month = MonthTitle;
                if (_scheduleItems.Count == 0 && _boardAll.Count == 0 && _reqAll.Count == 0 && _borrowAll.Count == 0)
                {
                    return $"{month} 还没有预领数据：先导入月排程，再每天导入 ORT 看板，开单后导入开单数据";
                }
                return $"{month}：月排程 {_scheduleItems.Count} 格 · 看板 {_boardAll.Count} 行 · 准备领用 {_reqAll.Count} 条"
                    + (_borrowAll.Count > 0 ? $" · 借用 {_borrowAll.Count} 条" : "");
            }
        }

        /* ###############################  命令  ################################ */

        private ICommand _refreshCommand;
        public ICommand RefreshCommand => _refreshCommand ??= new RelayCommand(ReloadAll);

        private ICommand _importScheduleCommand;
        public ICommand ImportScheduleCommand => _importScheduleCommand ??= new RelayCommand(ImportSchedule, () => CanImport);

        private ICommand _importBoardCommand;
        public ICommand ImportBoardCommand => _importBoardCommand ??= new RelayCommand(ImportBoard, () => CanImport);

        private ICommand _fetchBoardCommand;
        /// <summary>从公司平台直接取看板（登录 → 帆软看板 → 取数 → 入库），替代手工导出 Excel</summary>
        public ICommand FetchBoardCommand => _fetchBoardCommand ??= new RelayCommand(
            () => _ = FetchBoardFromPortalAsync(), () => CanImport && !FetchingBoard);

        /// <summary>正在从平台取数（按钮禁用 + 提示）</summary>
        public bool FetchingBoard
        {
            get => _fetchingBoard;
            private set
            {
                if (SetProperty(ref _fetchingBoard, value))
                {
                    OnPropertyChanged(nameof(FetchBoardButtonText));
                }
            }
        }
        private bool _fetchingBoard;

        /// <summary>取数按钮文字（取数中显示进度）</summary>
        public string FetchBoardButtonText
            => FetchingBoard ? "取数中…" : LanguageService.Get("PreIssue_FetchBoard");

        private ICommand _importRequisitionCommand;
        public ICommand ImportRequisitionCommand => _importRequisitionCommand ??= new RelayCommand(ImportRequisition, () => CanImport);

        private ICommand _syncPendingCommand;
        public ICommand SyncPendingCommand => _syncPendingCommand ??= new RelayCommand(SyncPending, () => CanImport);

        private ICommand _addBlacklistCommand;
        public ICommand AddBlacklistCommand => _addBlacklistCommand ??= new CommunityToolkit.Mvvm.Input.RelayCommand<object>(AddToBlacklist, p => CanImport && p is OrtBoardRow);

        private ICommand _undoImportCommand;
        public ICommand UndoImportCommand => _undoImportCommand ??= new CommunityToolkit.Mvvm.Input.RelayCommand<object>(UndoImport, p => CanImport);

        private ICommand _clearMonthCommand;
        public ICommand ClearMonthCommand => _clearMonthCommand ??= new CommunityToolkit.Mvvm.Input.RelayCommand<object>(ClearMonth, p => CanImport);

        private ICommand _prevWeekCommand;
        public ICommand PrevWeekCommand => _prevWeekCommand ??= new RelayCommand(PrevWeek);

        private ICommand _nextWeekCommand;
        public ICommand NextWeekCommand => _nextWeekCommand ??= new RelayCommand(NextWeek);

        private ICommand _currentWeekCommand;
        public ICommand CurrentWeekCommand => _currentWeekCommand ??= new RelayCommand(GoCurrentWeek);

        /// <summary>整体重载（月份切换/刷新/导入后）</summary>
        public void ReloadAll()
        {
            try
            {
                // 机种后来被加入黑名单（含正则）的，刷新时把它的待开单清掉——ORT 做不了的不该留在准备领用里
                PurgeBlacklistedPending();
                LoadSchedule();
                LoadBoard();
                LoadRequisition();
                LoadBorrow();
                OnPropertyChanged(nameof(HasAnyData));
                OnPropertyChanged(nameof(OverallHint));
                if (string.IsNullOrWhiteSpace(StatusMessage) || StatusMessage == "已刷新")
                {
                    StatusMessage = $"已刷新（{MonthTitle}）";
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"刷新预领数据失败: {ex.Message}");
                StatusMessage = "刷新失败：" + ex.Message;
            }
        }

        /// <summary>按黑名单清理本月的待开单（有清理时提示一条；已开单的不动）</summary>
        private void PurgeBlacklistedPending()
        {
            try
            {
                int removed = _service.PurgeBlacklistedPending(Month, out List<string> models);
                if (removed > 0)
                {
                    string sample = string.Join("、", models.Take(6)) + (models.Count > 6 ? "…" : "");
                    string message = $"已按黑名单清理 {removed} 个待开单：{sample}";
                    StatusMessage = message;
                    ToastService.Show(message);
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"按黑名单清理待开单失败: {ex.Message}");
            }
        }

        /* ###############################  准备领用：手工编辑  ################################ */

        /// <summary>能不能在表格里手工改准备领用（需要计划编辑权限）</summary>
        public bool CanEditRequisitionData => _permission.Can("plan.edit");

        /// <summary>表格里改完一格后保存这一行（失败时从库里把旧值读回来覆盖界面）</summary>
        public bool SaveReqRow(PreIssueRequisition row)
        {
            if (row == null)
            {
                return false;
            }
            if (!CanEditRequisitionData)
            {
                _ = MessageBox.Show(LanguageService.Get("PreIssue_ReqEditDenied"),
                    LanguageService.Get("PreIssue_RequisitionTab"), MessageBoxButton.OK, MessageBoxImage.Warning);
                ReloadRowFromDb(row);
                return false;
            }
            if (!_service.SaveRequisition(row, out string error, out int changed))
            {
                _ = MessageBox.Show(string.Format(LanguageService.Get("PreIssue_ReqSaveFailed"), error),
                    LanguageService.Get("PreIssue_RequisitionTab"), MessageBoxButton.OK, MessageBoxImage.Warning);
                ReloadRowFromDb(row);
                return false;
            }
            if (changed > 0)
            {
                StatusMessage = string.Format(LanguageService.Get("PreIssue_ReqSavedFormat"), row.PartNo, changed);
                // 阶段可能跟着开单号变了、汇总也要更新
                UpdateReqSummary();
                RefreshView(ReqView);
            }
            return true;
        }

        /// <summary>保存失败时把这一行恢复成库里的值（避免界面显示"改成功了"）</summary>
        private void ReloadRowFromDb(PreIssueRequisition row)
        {
            try
            {
                PreIssueRequisition fresh = _service.LoadRequisition(Month).FirstOrDefault(r => r.Id == row.Id);
                if (fresh == null)
                {
                    return;
                }
                row.Vchno = fresh.Vchno;
                row.WareCode = fresh.WareCode;
                row.TrCode = fresh.TrCode;
                row.PartNo = fresh.PartNo;
                row.ReqQty = fresh.ReqQty;
                row.IverId = fresh.IverId;
                row.VchStatus = fresh.VchStatus;
                row.IssueDept = fresh.IssueDept;
                row.CloseCode = fresh.CloseCode;
                row.Pc = fresh.Pc;
                row.State = fresh.State;
                RefreshView(ReqView);
            }
            catch (Exception ex)
            {
                _logger.Warn($"恢复准备领用旧值失败: {ex.Message}");
            }
        }

        /* ###############################  准备领用：列筛选 / 排序  ################################ */

        /// <summary>准备领用的列筛选（property → 勾选的值；多选，命中任意一个即显示）</summary>
        private readonly Dictionary<string, HashSet<string>> _reqColumnFilters = new(StringComparer.Ordinal);

        /// <summary>这一列有没有筛选</summary>
        public bool HasReqColumnFilter(string property) =>
            !string.IsNullOrEmpty(property) && _reqColumnFilters.ContainsKey(property);

        /// <summary>有没有任何列筛选</summary>
        public bool HasAnyReqColumnFilter => _reqColumnFilters.Count > 0;

        /// <summary>某一列当前勾选的值（没有筛选时返回 null = 全部显示）</summary>
        public HashSet<string> GetReqColumnSelection(string property) =>
            !string.IsNullOrEmpty(property) && _reqColumnFilters.TryGetValue(property, out HashSet<string> set)
                ? new HashSet<string>(set, StringComparer.Ordinal)
                : null;

        /// <summary>取一列的去重值（列头筛选菜单用）</summary>
        public List<string> GetReqColumnValues(string property)
        {
            System.Reflection.PropertyInfo info = typeof(PreIssueRequisition).GetProperty(property ?? "");
            if (info == null)
            {
                return [];
            }
            return _reqAll
                .Select(row => info.GetValue(row) as string ?? "")
                .Distinct(StringComparer.Ordinal)
                .OrderBy(v => v, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        /// <summary>
        /// 设置某一列的筛选：没勾 / 全部勾上 都等于「不筛选」（与领退表、计划表的口径一致）。
        /// 界面上的勾选是惰性的，点「确定」才走到这里。
        /// </summary>
        public void SetReqColumnFilter(string property, IEnumerable<string> values)
        {
            if (string.IsNullOrEmpty(property))
            {
                return;
            }
            HashSet<string> set = new(values ?? [], StringComparer.Ordinal);
            List<string> all = GetReqColumnValues(property);
            bool everything = all.Count == 0 || (set.Count >= all.Count && all.All(set.Contains));
            if (set.Count == 0 || everything)
            {
                _reqColumnFilters.Remove(property);
            }
            else
            {
                _reqColumnFilters[property] = set;
            }
            RefreshView(ReqView);
            UpdateReqSummary();
        }

        /// <summary>清除某一列 / 全部列筛选</summary>
        public void ClearReqColumnFilter(string property)
        {
            if (string.IsNullOrEmpty(property))
            {
                return;
            }
            _reqColumnFilters.Remove(property);
            RefreshView(ReqView);
            UpdateReqSummary();
        }

        /// <summary>清除全部列筛选</summary>
        public void ClearAllReqColumnFilters()
        {
            _reqColumnFilters.Clear();
            RefreshView(ReqView);
            UpdateReqSummary();
        }

        /// <summary>按列排序（direction 为空 = 不排序，恢复默认顺序）</summary>
        public void SortRequisitions(string property, ListSortDirection? direction)
        {
            try
            {
                using (ReqView.DeferRefresh())
                {
                    ReqView.SortDescriptions.Clear();
                    if (!string.IsNullOrEmpty(property) && direction.HasValue)
                    {
                        ReqView.SortDescriptions.Add(new SortDescription(property, direction.Value));
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"准备领用排序失败: {ex.Message}");
            }
        }

        /// <summary>当前排序（列头菜单打勾用）：没有排序返回 null</summary>
        public (string Property, ListSortDirection Direction)? CurrentReqSort
        {
            get
            {
                SortDescriptionCollection descriptions = ReqView.SortDescriptions;
                return descriptions.Count == 0 ? null : (descriptions[0].PropertyName, descriptions[0].Direction);
            }
        }

        /// <summary>重新算「可见 N 条」的汇总</summary>
        private void UpdateReqSummary()
        {
            int pending = _reqAll.Count(r => PreIssueCodes.IsPendingState(r.State) || string.IsNullOrWhiteSpace(r.Vchno));
            int opened = _reqAll.Count - pending;
            int visible = 0;
            foreach (object row in ReqView)
            {
                _ = row;
                visible++;
            }
            string baseText = _reqAll.Count == 0
                ? "没有记录"
                : $"{_reqAll.Count} 条 · 待开单 {pending} · 已开单 {opened}";
            ReqSummary = visible == _reqAll.Count
                ? baseText
                : baseText + $" · 可见 {visible} 条";
        }

        /* ###############################  借用表  ################################ */

        /// <summary>借用记录（全部月份，最新在前；状态筛选只管显示）</summary>
        public ObservableCollection<PreIssueBorrowItem> BorrowRows { get; } = [];

        public ICollectionView BorrowView { get; }

        private PreIssueBorrowItem _selectedBorrowRow;
        public PreIssueBorrowItem SelectedBorrowRow
        {
            get => _selectedBorrowRow;
            set => SetProperty(ref _selectedBorrowRow, value);
        }

        /// <summary>状态筛选（全部 / 待借用 / 借用中 / 已归还）</summary>
        public ObservableCollection<string> BorrowStatusOptions { get; } = [];

        private string _borrowStatusFilter;
        public string BorrowStatusFilter
        {
            get => _borrowStatusFilter;
            set
            {
                if (SetProperty(ref _borrowStatusFilter, value))
                {
                    RefreshView(BorrowView);
                    UpdateBorrowSummary();
                }
            }
        }

        private string _borrowSummary;
        public string BorrowSummary { get => _borrowSummary; private set => SetProperty(ref _borrowSummary, value); }

        private List<PreIssueBorrowItem> _borrowAll = [];

        private void LoadBorrow()
        {
            if (BorrowStatusOptions.Count == 0)
            {
                BorrowStatusOptions.Add("全部");
                BorrowStatusOptions.Add(PreIssueCodes.BorrowPending);
                BorrowStatusOptions.Add(PreIssueCodes.Borrowing);
                BorrowStatusOptions.Add(PreIssueCodes.BorrowReturned);
                _borrowStatusFilter = "全部";
            }
            try
            {
                _borrowAll = _service.LoadBorrowItems();
            }
            catch (Exception ex)
            {
                _logger.Warn($"读取借用表失败: {ex.Message}");
                _borrowAll = [];
            }
            BorrowRows.Clear();
            foreach (PreIssueBorrowItem row in _borrowAll)
            {
                BorrowRows.Add(row);
            }
            OnPropertyChanged(nameof(BorrowStatusFilter));
            RefreshView(BorrowView);
            UpdateBorrowSummary();
        }

        private bool FilterBorrow(object obj)
        {
            if (obj is not PreIssueBorrowItem row)
            {
                return false;
            }
            string filter = BorrowStatusFilter;
            if (string.IsNullOrEmpty(filter) || filter == "全部")
            {
                return true;
            }
            return string.Equals(StatusOf(row), filter, StringComparison.Ordinal);
        }

        private static string StatusOf(PreIssueBorrowItem row) =>
            string.IsNullOrWhiteSpace(row?.Status) ? PreIssueCodes.BorrowPending : row.Status.Trim();

        private void UpdateBorrowSummary()
        {
            if (_borrowAll.Count == 0)
            {
                BorrowSummary = "没有借用记录";
                return;
            }
            int pending = _borrowAll.Count(r => StatusOf(r) == PreIssueCodes.BorrowPending);
            int borrowing = _borrowAll.Count(r => StatusOf(r) == PreIssueCodes.Borrowing);
            int returned = _borrowAll.Count(r => StatusOf(r) == PreIssueCodes.BorrowReturned);
            int visible = 0;
            foreach (object row in BorrowView)
            {
                _ = row;
                visible++;
            }
            string text = $"{_borrowAll.Count} 条 · 待借用 {pending} · 借用中 {borrowing} · 已归还 {returned}";
            BorrowSummary = visible == _borrowAll.Count ? text : text + $" · 可见 {visible} 条";
        }

        /// <summary>能不能改借用表（需要计划编辑权限）</summary>
        public bool CanEditBorrowData => _permission.Can("plan.edit");

        /// <summary>借用表里改完一格后保存这一行（失败时从库里读回旧值）</summary>
        public bool SaveBorrowRow(PreIssueBorrowItem row)
        {
            if (row == null)
            {
                return false;
            }
            if (!CanEditBorrowData)
            {
                _ = MessageBox.Show(LanguageService.Get("PreIssue_ReqEditDenied"),
                    LanguageService.Get("PreIssue_BorrowTab"), MessageBoxButton.OK, MessageBoxImage.Warning);
                ReloadBorrowRow(row);
                return false;
            }
            if (!_service.SaveBorrowItem(row, out string error))
            {
                _ = MessageBox.Show($"保存借用记录失败：{error}", LanguageService.Get("PreIssue_BorrowTab"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                ReloadBorrowRow(row);
                return false;
            }
            // 服务端会把状态/机种/数量规整一遍（就地写回同一个对象），这里刷新汇总即可
            UpdateBorrowSummary();
            RefreshView(BorrowView);
            StatusMessage = $"已保存借用记录：{row.Model}";
            return true;
        }

        /// <summary>
        /// 保存失败时把这一行恢复成库里的值。
        /// 只改这一行的字段、不重建集合：单元格还处在编辑事务里，换整个 ItemsSource 会抛
        /// 「在 AddNew 或 EditItem 事务过程中不允许 …」。
        /// </summary>
        private void ReloadBorrowRow(PreIssueBorrowItem row)
        {
            try
            {
                PreIssueBorrowItem fresh = _service.LoadBorrowItems().FirstOrDefault(r => r.Id == row.Id);
                if (fresh == null)
                {
                    return;
                }
                row.Status = fresh.Status;
                row.Model = fresh.Model;
                row.Qty = fresh.Qty;
                row.BorrowDate = fresh.BorrowDate;
                row.ReturnDate = fresh.ReturnDate;
                RefreshView(BorrowView);
            }
            catch (Exception ex)
            {
                _logger.Warn($"恢复借用记录旧值失败: {ex.Message}");
            }
        }

        /// <summary>右键「登记借用」：状态 → 借用中，借用日期填今天</summary>
        public bool MarkBorrowedRow(PreIssueBorrowItem row) => ChangeBorrowState(row, true);

        /// <summary>右键「登记归还」：状态 → 已归还，归还日期填今天</summary>
        public bool MarkReturnedRow(PreIssueBorrowItem row) => ChangeBorrowState(row, false);

        private bool ChangeBorrowState(PreIssueBorrowItem row, bool borrowing)
        {
            if (row == null)
            {
                return false;
            }
            if (!CanEditBorrowData)
            {
                _ = MessageBox.Show(LanguageService.Get("PreIssue_ReqEditDenied"),
                    LanguageService.Get("PreIssue_BorrowTab"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
            string action = borrowing ? LanguageService.Get("PreIssue_Borrow_Menu_Borrow") : LanguageService.Get("PreIssue_Borrow_Menu_Return");
            bool ok = borrowing
                ? _service.MarkBorrowed(row.Id, out string error)
                : _service.MarkReturned(row.Id, out error);
            if (!ok)
            {
                _ = MessageBox.Show(error, action, MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
            ReloadBorrowRow(row);
            UpdateBorrowSummary();
            StatusMessage = $"{action}：{row.Model}";
            ToastService.Show(StatusMessage);
            return true;
        }

        /* ###############################  同系列客户优先级  ################################ */

        /// <summary>当前生效的客户优先级（数字越小越优先，可并列）</summary>
        public List<CustomerPriorityEntry> LoadCustomerPriority() => _service.LoadCustomerPriority();

        /// <summary>客户代码 → 客户名称（主数据库的客户代码表）</summary>
        public Dictionary<string, string> LoadCustomerCodeNames() => _service.LoadCustomerCodeNames();

        /// <summary>优先级编辑窗口的候选客户代码</summary>
        public List<string> LoadCustomerCodeCandidates() => _service.LoadCustomerCodeCandidates();

        /// <summary>保存客户优先级（存数据库，多台电脑共用）</summary>
        public void SaveCustomerPriority(IEnumerable<CustomerPriorityEntry> entries)
        {
            List<CustomerPriorityEntry> saved = PreIssueSeriesPriority.Normalize(entries);
            _service.SaveCustomerPriority(saved);
            StatusMessage = "客户优先级已保存：" + PreIssueSeriesPriority.Describe(saved);
        }

        /* ###############################  黑名单：编辑框辅助  ################################ */

        /// <summary>把单个机种拆成「同系列」正则的预览（编辑窗口用）</summary>
        public static bool TryBuildSeriesPattern(string model, out string pattern, out string series,
            out string customer, out string version) =>
            OrtBlacklistMatcher.TryBuildSeriesPattern(model, out pattern, out series, out customer, out version);

        /// <summary>一条黑名单规则本月会命中哪些机种（编辑窗口确认前预览）</summary>
        public List<string> PreviewBlacklistRule(string value, bool isPattern) =>
            _service.PreviewBlacklistRule(value, isPattern);

        private void PrevWeek() => StepWeek(-1);

        private void NextWeek()
        {
            if (WholeMonth)
            {
                WholeMonth = false;
            }
            StepWeek(1);
        }

        private void GoCurrentWeek()
        {
            WholeMonth = false;
            if (PickCurrentWeek() is WeekOption week)
            {
                SelectedWeek = week;
            }
        }

        private void StepWeek(int delta)
        {
            if (Weeks.Count == 0)
            {
                return;
            }
            int index = Weeks.IndexOf(SelectedWeek);
            if (index < 0)
            {
                index = 0;
            }
            int target = index + delta;
            if (target < 0)
            {
                target = 0;
            }
            if (target >= Weeks.Count)
            {
                target = Weeks.Count - 1;
            }
            SelectedWeek = Weeks[target];
        }

        private void ImportSchedule()
        {
            string file = _pathService.OpenPathDialog(LanguageService.Get("PreIssue_ImportScheduleTitle"),
                "Excel文件|*.xls;*.xlsx;*.xlsm|所有文件|*.*");
            if (string.IsNullOrWhiteSpace(file))
            {
                return;
            }
            // 先预览再确认：文件形状不对（比如把开单数据选成月排程）在这一步就被拦下来
            if (!ConfirmImport(_service.PreviewSchedule(file), LanguageService.Get("PreIssue_ScheduleTab")))
            {
                return;
            }
            RunImport(() => _service.ImportSchedule(file), LanguageService.Get("PreIssue_ScheduleTab"));
        }

        private void ImportBoard()
        {
            string file = _pathService.OpenPathDialog(LanguageService.Get("PreIssue_ImportBoardTitle"),
                "Excel文件|*.xls;*.xlsx;*.xlsm|所有文件|*.*");
            if (string.IsNullOrWhiteSpace(file))
            {
                return;
            }
            if (!ConfirmImport(_service.PreviewBoard(file), LanguageService.Get("PreIssue_BoardTab")))
            {
                return;
            }
            RunImport(() => _service.ImportBoard(file), LanguageService.Get("PreIssue_BoardTab"));
        }

        /// <summary>
        /// 从公司平台直接取 ORT 看板：登录（必要时用保存的密码重登）→ 打开目标看板 → 取数 → 认表头还原成行
        /// → 先预览确认（和 Excel 导入一样的确认框）→ 走同一套入库/变动记录。
        /// 取数条件（日期区间、MoType）与链接配置在「设置 → ORT看板取数」里维护。
        /// </summary>
        private async Task FetchBoardFromPortalAsync()
        {
            string label = LanguageService.Get("PreIssue_BoardTab");
            if (FetchingBoard)
            {
                return;
            }
            FetchingBoard = true;
            StatusMessage = "正在从公司平台取看板…";
            try
            {
                OrtBoardFetchResult result = await _board.FetchAsync().ConfigureAwait(true);
                if (!result.Success)
                {
                    StatusMessage = "从平台取数失败：" + result.Message;
                    _ = MessageBox.Show(result.Message ?? "取数失败", label, MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }
                StatusMessage = $"已取到 {result.Rows.Count} 行（{result.QueryText}）";
                PreIssueImportPreview preview = _board.Preview(result.Rows,
                    $"平台取数（{result.Widget}）");
                if (!ConfirmImport(preview, label))
                {
                    StatusMessage = "已取消导入（取到的数据没有写库）";
                    return;
                }
                string source = $"平台取数 {DateTime.Now:yyyy-MM-dd HH:mm}";
                RunImport(() => _board.Import(result.Rows, source), label);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "从平台取看板失败");
                StatusMessage = "从平台取数失败：" + ex.Message;
                _ = MessageBox.Show(ex.Message, label, MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                FetchingBoard = false;
            }
        }

        private void ImportRequisition()
        {
            string file = _pathService.OpenPathDialog(LanguageService.Get("PreIssue_ImportRequisitionTitle"),
                "Excel文件|*.xls;*.xlsx;*.xlsm|所有文件|*.*");
            if (string.IsNullOrWhiteSpace(file))
            {
                return;
            }
            string month = Month;
            if (!ConfirmImport(_service.PreviewRequisition(file, month), LanguageService.Get("PreIssue_RequisitionTab")))
            {
                return;
            }
            RunImport(() => _service.ImportRequisition(file, month), LanguageService.Get("PreIssue_RequisitionTab"));
        }

        /// <summary>
        /// 导入前的确认：文件不像本页该导的数据就直接拦下（不写库），
        /// 像的话把「哪个月、多少条、会覆盖多少」摆出来让用户确认。
        /// </summary>
        private bool ConfirmImport(PreIssueImportPreview preview, string label)
        {
            if (preview == null)
            {
                return false;
            }
            if (!preview.Success)
            {
                StatusMessage = $"{label}没有导入：{preview.Message}";
                _ = MessageBox.Show(preview.Message, label, MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
            return MessageBox.Show(preview.Describe(), label, MessageBoxButton.YesNo, MessageBoxImage.Question)
                == MessageBoxResult.Yes;
        }

        /// <summary>撤销上次导入（参数：schedule / board）</summary>
        private void UndoImport(object parameter)
        {
            string source = SourceOfParameter(parameter);
            if (source == null)
            {
                return;
            }
            string text = string.Format(LanguageService.Get("PreIssue_UndoConfirm"), MonthTitle, source);
            if (MessageBox.Show(text, LanguageService.Get("PreIssue_UndoTitle"),
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            {
                return;
            }
            try
            {
                PreIssueImportResult result = _service.UndoImport(source, Month);
                ReloadAll();
                if (!result.Success)
                {
                    StatusMessage = result.Message;
                    _ = MessageBox.Show(result.Message, LanguageService.Get("PreIssue_UndoTitle"),
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                StatusMessage = result.Message;
                _ = MessageBox.Show(result.Message, LanguageService.Get("PreIssue_UndoTitle"),
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                _logger.Warn($"撤销导入失败: {ex.Message}");
                StatusMessage = "撤销失败：" + ex.Message;
            }
        }

        /// <summary>清除当前月份某一类数据（参数：schedule / board）</summary>
        private void ClearMonth(object parameter)
        {
            string source = SourceOfParameter(parameter);
            if (source == null)
            {
                return;
            }
            int count = source == PreIssueCodes.SourceSchedule ? _scheduleItems.Count : _boardAll.Count;
            string text = string.Format(LanguageService.Get("PreIssue_ClearConfirm"), MonthTitle, source, count);
            if (MessageBox.Show(text, LanguageService.Get("PreIssue_ClearTitle"),
                MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes)
            {
                return;
            }
            try
            {
                PreIssueImportResult result = _service.ClearMonth(source, Month);
                ReloadAll();
                StatusMessage = result.Message;
                _ = MessageBox.Show(result.Message, LanguageService.Get("PreIssue_ClearTitle"),
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                _logger.Warn($"清除月份数据失败: {ex.Message}");
                StatusMessage = "清除失败：" + ex.Message;
            }
        }

        /// <summary>按钮参数（schedule / board）→ 变动记录里的来源名</summary>
        private static string SourceOfParameter(object parameter)
        {
            if (parameter is not string key)
            {
                return null;
            }
            return key switch
            {
                "schedule" => PreIssueCodes.SourceSchedule,
                "board" => PreIssueCodes.SourceBoard,
                _ => null
            };
        }

        /// <summary>执行导入并刷新界面（统一处理异常与结果提示）</summary>
        private void RunImport(Func<PreIssueImportResult> import, string label)
        {
            try
            {
                PreIssueImportResult result = import();
                if (!result.Success)
                {
                    StatusMessage = $"{label}导入失败：{result.Message}";
                    _ = MessageBox.Show(string.IsNullOrWhiteSpace(result.Message) ? "导入失败" : result.Message,
                        label, MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                string keepMonth = result.Month;
                LoadMonths();
                if (!string.IsNullOrWhiteSpace(keepMonth) && SelectedMonth?.Month != keepMonth
                    && Months.FirstOrDefault(m => m.Month == keepMonth) is MonthOption option)
                {
                    SelectedMonth = option;
                }
                else
                {
                    ReloadAll();
                }
                StatusMessage = $"{label}导入完成：{result.Summary()}";
                _ = MessageBox.Show($"{result.Summary()}\n\n变动已记入「变动记录」，可在工具栏查看。",
                    label, MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{label}导入失败");
                StatusMessage = $"{label}导入失败：{ex.Message}";
                _ = MessageBox.Show(ex.Message, label, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SyncPending()
        {
            try
            {
                int created = _service.SyncPendingFromBoard(Month, out List<string> models);
                ReloadAll();
                StatusMessage = created == 0
                    ? "看板里没有新的已入库机种（黑名单和已有记录都跳过了）"
                    : $"从看板补全待开单 {created} 个：{string.Join("、", models.Take(12))}{(models.Count > 12 ? "…" : "")}";
            }
            catch (Exception ex)
            {
                _logger.Warn($"从看板补全待开单失败: {ex.Message}");
                StatusMessage = "补全失败：" + ex.Message;
            }
        }

        /// <summary>
        /// 手工添加一个待开单机种（窗口由界面层弹出，拿到机种 / 工令 / 数量 / 备注后调这里）。
        /// 同机种本月已有记录时**允许再加**，但先提醒一句：重复领用建议在备注里写明原因。
        /// </summary>
        public bool AddPendingModel(string model, string jobNo, string qty = null, string note = null)
        {
            int existing = _service.CountRequisitionsByModel(Month, model);
            if (existing > 0)
            {
                string ask = $"本月已经有 {existing} 条「{model}」的记录，还要再添加一条吗？\n\n"
                    + "同机种重复领用，建议在「备注」里写明原因（备注会记进变动记录）。";
                if (MessageBox.Show(ask, LanguageService.Get("PreIssue_AddPendingTitle"),
                    MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                {
                    return false;
                }
            }
            if (!_service.AddPending(Month, model, jobNo, qty, note, out string error))
            {
                _ = MessageBox.Show(error, LanguageService.Get("PreIssue_AddPendingTitle"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
            ReloadAll();
            string message = $"已添加待开单：{model}";
            if (!string.IsNullOrWhiteSpace(qty))
            {
                message += $"（数量 {qty}）";
            }
            if (existing > 0)
            {
                message += $"；本月同机种第 {existing + 1} 条";
            }
            StatusMessage = message;
            return true;
        }

        /// <summary>看板里出现过的机种（添加待开单时的候选，方便直接挑）</summary>
        public List<string> BoardModelCandidates() => _boardAll
            .Where(b => !string.IsNullOrWhiteSpace(b.Model))
            .Select(b => b.Model.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(m => m, StringComparer.OrdinalIgnoreCase)
            .ToList();

        /// <summary>能不能删（只有「待开单」且没有开单号的记录可删）</summary>
        public bool CanDeletePending(PreIssueRequisition row) =>
            row != null
            && PreIssueCodes.IsPendingState(row.State)
            && string.IsNullOrWhiteSpace(row.Vchno);

        /// <summary>
        /// 删除待开单：确认框由界面层弹（三个按钮——是 / 是并加入黑名单 / 否），
        /// addBlacklist = true 时删除后再把机种加入黑名单（以后看板变动不再推进它）。
        /// </summary>
        public void DeletePending(PreIssueRequisition row, bool addBlacklist)
        {
            if (row == null)
            {
                return;
            }
            if (!CanDeletePending(row))
            {
                _ = MessageBox.Show(LanguageService.Get("PreIssue_OnlyPendingDeletable"),
                    LanguageService.Get("PreIssue_DeletePendingTitle"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            string part = row.PartNo ?? "";
            if (!_service.DeletePending(row.Id, out string error))
            {
                _ = MessageBox.Show(error, LanguageService.Get("PreIssue_DeletePendingTitle"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            string message = $"已删除待开单：{part}";
            if (addBlacklist)
            {
                if (_service.AddToBlacklist(part, LanguageService.Get("PreIssue_BlacklistFromDelete"), false, out string blackError))
                {
                    message += "，并已加入黑名单（以后不再推进到准备领用）";
                }
                else
                {
                    message += $"；加入黑名单失败：{blackError}";
                }
            }
            ReloadAll();
            StatusMessage = message;
            ToastService.Show(message);
        }

        private void AddToBlacklist(object parameter)
        {
            string model = parameter switch
            {
                OrtBoardRow row => row.Model,
                PreIssueRequisition req => req.PartNo,
                string text => text,
                _ => null
            };
            if (string.IsNullOrWhiteSpace(model))
            {
                return;
            }
            if (!_service.AddToBlacklist(model, null, out string error))
            {
                _ = MessageBox.Show(error, LanguageService.Get("PreIssue_AddBlacklistTitle"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            ReloadAll();
            StatusMessage = $"已加入黑名单：{model}（该机种不再转准备领用）";
        }

        /* ###############################  小工具  ################################ */

        private static bool Contains(string source, string keyword) =>
            !string.IsNullOrEmpty(source) && source.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0;

        private static void RefreshView(ICollectionView view)
        {
            try
            {
                view?.Refresh();
            }
            catch (Exception)
            {
                // 编辑事务中刷新会抛异常：延后一次再试，避免打断用户操作
                Application.Current?.Dispatcher?.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        view?.Refresh();
                    }
                    catch
                    {
                        // 忽略：下一次重载会重新构建
                    }
                }));
            }
        }
    }

    /// <summary>月份选项</summary>
    public class MonthOption
    {
        /// <summary>yyyyMM</summary>
        public string Month { get; set; }

        /// <summary>显示文本（2026/11）</summary>
        public string Text { get; set; }

        public override string ToString() => Text;
    }

    /// <summary>周视图的一周</summary>
    public class WeekOption
    {
        public int Index { get; set; }

        /// <summary>这周的周一</summary>
        public DateTime Start { get; set; }

        /// <summary>显示文本（11/2～11/8）</summary>
        public string Text { get; set; }

        public override string ToString() => Text;
    }

    /// <summary>排程矩阵的一个日列</summary>
    public class ScheduleDayColumn
    {
        /// <summary>日（1..31；不在本月时为 0，列显示「-」）</summary>
        public int Day { get; set; }

        /// <summary>列头（4日 / 六）</summary>
        public string Header { get; set; }

        /// <summary>是否周末（列头高亮）</summary>
        public bool IsWeekend { get; set; }

        /// <summary>这一列是否属于当前月份</summary>
        public bool InMonth { get; set; }
    }

    /// <summary>排程矩阵的一行（線別 + 機種）</summary>
    public class PreScheduleRow
    {
        public string Line { get; set; }

        public string Model { get; set; }

        /// <summary>每天的数量文本（下标 = 日，0 为空）</summary>
        public string[] Days { get; set; } = new string[32];

        /// <summary>每天的数量（排序/统计用）</summary>
        public int[] Quantities { get; set; } = new int[32];

        /// <summary>本月合计</summary>
        public int Total { get; set; }

        public string LineModel => $"{Line} / {Model}";
    }
}
