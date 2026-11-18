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

        public PreIssueViewModel(PreIssueService service, IPermissionService permission, IPathService pathService)
        {
            _service = service;
            _permission = permission;
            _pathService = pathService;
            BoardView = CollectionViewSource.GetDefaultView(BoardRows);
            BoardView.Filter = FilterBoard;
            ReqView = CollectionViewSource.GetDefaultView(ReqRows);
            ReqView.Filter = FilterRequisition;
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
            BoardSummary = total == 0
                ? "没有看板数据"
                : $"{total} 行 · 已入库 {stocked} · 待转领用 {pending} · SIM94禁止出货 {forbidden}";
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

        private string _reqHint;
        public string ReqHint { get => _reqHint; private set => SetProperty(ref _reqHint, value); }

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
            int pending = _reqAll.Count(r => PreIssueCodes.IsPendingState(r.State) || string.IsNullOrWhiteSpace(r.Vchno));
            int opened = _reqAll.Count - pending;
            ReqSummary = _reqAll.Count == 0
                ? "没有记录"
                : $"{_reqAll.Count} 条 · 待开单 {pending} · 已开单 {opened}";
            ReqHint = pending > 0
                ? $"有 {pending} 个机种待开单：开好单后把下载的开单数据（11.XLS 一类）导入进来补全"
                : "待开单都开好了：可在开单流程有变化时再次导入开单数据，同步 VCH_STATUS / CLOSE_CODE";
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

        /// <summary>加入黑名单（黑名单管理窗口用）</summary>
        public bool AddBlacklistModel(string model, string reason) =>
            _service.AddToBlacklist(model, reason, out string error) || Fail(error);

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
        public bool HasAnyData => _scheduleItems.Count > 0 || _boardAll.Count > 0 || _reqAll.Count > 0;

        /// <summary>总体提示（预领页签顶部的一行说明）</summary>
        public string OverallHint
        {
            get
            {
                string month = MonthTitle;
                return _scheduleItems.Count == 0 && _boardAll.Count == 0 && _reqAll.Count == 0
                    ? $"{month} 还没有预领数据：先导入月排程，再每天导入 ORT 看板，开单后导入开单数据"
                    : $"{month}：月排程 {_scheduleItems.Count} 格 · 看板 {_boardAll.Count} 行 · 准备领用 {_reqAll.Count} 条";
            }
        }

        /* ###############################  命令  ################################ */

        private ICommand _refreshCommand;
        public ICommand RefreshCommand => _refreshCommand ??= new RelayCommand(ReloadAll);

        private ICommand _importScheduleCommand;
        public ICommand ImportScheduleCommand => _importScheduleCommand ??= new RelayCommand(ImportSchedule, () => CanImport);

        private ICommand _importBoardCommand;
        public ICommand ImportBoardCommand => _importBoardCommand ??= new RelayCommand(ImportBoard, () => CanImport);

        private ICommand _importRequisitionCommand;
        public ICommand ImportRequisitionCommand => _importRequisitionCommand ??= new RelayCommand(ImportRequisition, () => CanImport);

        private ICommand _syncPendingCommand;
        public ICommand SyncPendingCommand => _syncPendingCommand ??= new RelayCommand(SyncPending, () => CanImport);

        private ICommand _deletePendingCommand;
        public ICommand DeletePendingCommand => _deletePendingCommand ??= new CommunityToolkit.Mvvm.Input.RelayCommand<object>(DeletePending, p => CanImport && p is PreIssueRequisition);

        private ICommand _addBlacklistCommand;
        public ICommand AddBlacklistCommand => _addBlacklistCommand ??= new CommunityToolkit.Mvvm.Input.RelayCommand<object>(AddToBlacklist, p => CanImport && p is OrtBoardRow);

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
                LoadSchedule();
                LoadBoard();
                LoadRequisition();
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
            RunImport(() => _service.ImportSchedule(file), "月排程");
        }

        private void ImportBoard()
        {
            string file = _pathService.OpenPathDialog(LanguageService.Get("PreIssue_ImportBoardTitle"),
                "Excel文件|*.xls;*.xlsx;*.xlsm|所有文件|*.*");
            if (string.IsNullOrWhiteSpace(file))
            {
                return;
            }
            RunImport(() => _service.ImportBoard(file), "ORT看板");
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
            RunImport(() => _service.ImportRequisition(file, month), "准备领用");
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

        /// <summary>手工添加一个待开单机种（窗口由界面层弹出，拿到机种后调这里）</summary>
        public bool AddPendingModel(string model, string jobNo)
        {
            if (!_service.AddPending(Month, model, jobNo, out string error))
            {
                _ = MessageBox.Show(error, LanguageService.Get("PreIssue_AddPendingTitle"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
            ReloadAll();
            StatusMessage = $"已添加待开单：{model}";
            return true;
        }

        /// <summary>看板里出现过的机种（添加待开单时的候选，方便直接挑）</summary>
        public List<string> BoardModelCandidates() => _boardAll
            .Where(b => !string.IsNullOrWhiteSpace(b.Model))
            .Select(b => b.Model.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(m => m, StringComparer.OrdinalIgnoreCase)
            .ToList();

        private void DeletePending(object parameter)
        {
            if (parameter is not PreIssueRequisition row)
            {
                return;
            }
            if (!_service.DeletePending(row.Id, out string error))
            {
                _ = MessageBox.Show(error, LanguageService.Get("PreIssue_DeletePendingTitle"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            ReloadAll();
            StatusMessage = $"已删除待开单：{row.PartNo}";
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
