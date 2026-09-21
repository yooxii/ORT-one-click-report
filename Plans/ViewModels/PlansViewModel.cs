using CommunityToolkit.Mvvm.ComponentModel;
using Newtonsoft.Json;
using NLog;
using ORT一键报告.Models;
using ORT一键报告.Services;
using ORT一键报告.Utils;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Threading;

namespace ORT一键报告.Plans.ViewModels
{
    /// <summary>
    /// 单列筛选条件：列属性名 + 列标题 + 该列全部可选值 + 已选值。
    /// 已选值为空表示"全部"（不过滤）；值统一按表格显示文本存储，日期为 yyyy/M/d。
    /// </summary>
    public class ColumnFilter
    {
        /// <summary>列绑定的属性名（如 ModelName）</summary>
        public string Property { get; }

        /// <summary>列标题（筛选菜单显示用，已本地化）</summary>
        public string Label { get; set; }

        /// <summary>该列全部可选值（去重并排序）</summary>
        public List<string> Options { get; } = [];

        /// <summary>已选中的值（空集合表示不过滤）</summary>
        public HashSet<string> Selected { get; } = new(StringComparer.Ordinal);

        /// <summary>该列是否有生效的筛选</summary>
        public bool IsActive => Selected.Count > 0;

        public ColumnFilter(string property, string label)
        {
            Property = property;
            Label = label;
        }
    }

    /// <summary>
    /// 领退和计划主界面 ViewModel：领退表与计划表分表展示（两个 Tab），
    /// 暂存修改 + 手动提交 + 变更日志，自动补全（D/C、線別、回线RT工令、工作编号等）。
    /// </summary>
    public partial class PlansViewModel : ObservableObject
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;
        private readonly PlanExcelService _excelService;
        private readonly IPathService _pathService;
        private readonly IPermissionService _permission;
        private readonly ReviewService _reviewService;
        private readonly AdminService _adminService;
        private readonly AppSettingsService _appSettings;
        private readonly ReportScanService _reportScan;
        private readonly HighCostTaskCoordinator _coordinator;

        /// <summary>
        /// 报告链接缓存：工作编号 → 报告夹信息（扫描自报告路径）
        /// </summary>
        private readonly Dictionary<string, ReportLink> _reportLinks = [];

        /// <summary>
        /// 计划表列筛选：列属性名 → 筛选条件（每列最多一个条件，互相独立）
        /// </summary>
        private readonly Dictionary<string, ColumnFilter> _planFilters = new(StringComparer.Ordinal);

        /// <summary>
        /// 领退表列筛选：列属性名 → 筛选条件（与计划表互不影响）
        /// </summary>
        private readonly Dictionary<string, ColumnFilter> _reqFilters = new(StringComparer.Ordinal);

        /// <summary>
        /// 属性取值缓存：(类型, 属性名) → PropertyInfo
        /// </summary>
        private static readonly Dictionary<(Type, string), System.Reflection.PropertyInfo> PropertyCache = [];

        /* ###############################  领退表集合  ################################ */

        /// <summary>
        /// 领退表全部记录
        /// </summary>
        public ObservableCollection<Requisition> Requisitions { get; } = [];

        /// <summary>
        /// 领退表带筛选的视图
        /// </summary>
        public ICollectionView RequisitionsView { get; }

        private Requisition _selectedRequisition;
        /// <summary>
        /// 当前选中的领退表记录
        /// </summary>
        public Requisition SelectedRequisition { get => _selectedRequisition; set => SetProperty(ref _selectedRequisition, value); }

        /* ###############################  计划表集合  ################################ */

        /// <summary>
        /// 计划表全部记录
        /// </summary>
        public ObservableCollection<Plan> Plans { get; } = [];

        /// <summary>
        /// 计划表带筛选的视图
        /// </summary>
        public ICollectionView PlansView { get; }

        private Plan _selectedPlan;
        /// <summary>
        /// 当前选中的计划表记录
        /// </summary>
        public Plan SelectedPlan { get => _selectedPlan; set => SetProperty(ref _selectedPlan, value); }

        /// <summary>
        /// 计划表当前生效的列筛选（供界面显示筛选状态）
        /// </summary>
        public IReadOnlyDictionary<string, ColumnFilter> PlanFilters => _planFilters;

        /* ---------- 计划表状况统计（搜索栏右侧显示，随搜索与列筛选变化） ---------- */

        private int _ongoingCount;
        /// <summary>当前可见（已过搜索与列筛选）的计划里「测试中」的条数</summary>
        public int OngoingCount { get => _ongoingCount; private set => SetProperty(ref _ongoingCount, value); }

        private int _pendingCount;
        /// <summary>当前可见的计划里「预排测试」的条数</summary>
        public int PendingCount { get => _pendingCount; private set => SetProperty(ref _pendingCount, value); }

        private int _closedCount;
        /// <summary>当前可见的计划里「已结案」的条数</summary>
        public int ClosedCount { get => _closedCount; private set => SetProperty(ref _closedCount, value); }

        /// <summary>
        /// 刷新状况统计：只统计当前搜索与列筛选后仍然可见的计划（与表格所见一致）
        /// </summary>
        public void UpdateStatusCounts()
        {
            int ongoing = 0, pending = 0, closed = 0;
            foreach (Plan plan in PlansView.OfType<Plan>())
            {
                switch (PlanStatusKind.Of(plan.Status))
                {
                    case PlanStatusKind.Ongoing:
                        ongoing++;
                        break;
                    case PlanStatusKind.Pending:
                        pending++;
                        break;
                    case PlanStatusKind.Closed:
                        closed++;
                        break;
                }
            }
            OngoingCount = ongoing;
            PendingCount = pending;
            ClosedCount = closed;
        }

        /// <summary>
        /// 领退表当前生效的列筛选
        /// </summary>
        public IReadOnlyDictionary<string, ColumnFilter> ReqFilters => _reqFilters;

        private string _statusMessage = "就绪";
        /// <summary>
        /// 状态栏消息
        /// </summary>
        public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }

        /// <summary>
        /// 导入/导出/清空权限（技术员及以上）
        /// </summary>
        public bool CanImportExport => _permission.Can("plan.import");

        /// <summary>
        /// 新增/编辑/删除权限（普通用户及以上；普通用户提交审核）
        /// </summary>
        public bool CanEdit => _permission.Can("plan.edit");

        /// <summary>
        /// 当前编辑是否需提交审核（普通用户）
        /// </summary>
        public bool NeedsReview => _permission.PlanEditNeedsReview;

        /// <summary>
        /// 是否允许表格内直接编辑（技术员及以上）
        /// </summary>
        public bool CanGridEdit => CanEdit && !NeedsReview;

        /// <summary>
        /// 表格是否只读（XAML 绑定用）
        /// </summary>
        public bool IsGridReadOnly => !CanGridEdit;

        /* ###############################  暂存修改  ################################ */

        private readonly Dictionary<long, Plan> _planOriginals = [];
        private readonly Dictionary<long, Requisition> _reqOriginals = [];
        private readonly List<Plan> _pendingPlanAdded = [];
        private readonly List<Requisition> _pendingReqAdded = [];
        private readonly Dictionary<long, Plan> _pendingPlanDeleted = [];
        private readonly Dictionary<long, Requisition> _pendingReqDeleted = [];

        /// <summary>
        /// 是否有未提交的修改
        /// </summary>
        public bool HasPendingChanges => _pendingPlanAdded.Count > 0 || _pendingReqAdded.Count > 0
            || _pendingPlanDeleted.Count > 0 || _pendingReqDeleted.Count > 0
            || DetectPlanModifiedCount() > 0 || DetectReqModifiedCount() > 0;

        /// <summary>
        /// 未提交修改的描述（状态栏显示）
        /// </summary>
        public string PendingText => HasPendingChanges
            ? string.Format(LanguageService.Get("Plans_PendingFormat"), _pendingReqAdded.Count, _pendingPlanAdded.Count, DetectPlanModifiedCount() + DetectReqModifiedCount(), _pendingPlanDeleted.Count + _pendingReqDeleted.Count)
            : LanguageService.Get("Plans_NoPending");

        /* ###############################  字典  ################################ */

        /// <summary>
        /// 测试项目字典
        /// </summary>
        public List<string> CatalogTestItems { get; private set; } = [];

        /// <summary>
        /// 产品别字典
        /// </summary>
        public List<string> CatalogProducts { get; private set; } = [];

        /// <summary>
        /// 客户别字典
        /// </summary>
        public List<string> CatalogCustomers { get; private set; } = [];

        /// <summary>
        /// 阶段字典
        /// </summary>
        public List<string> CatalogStages { get; private set; } = [];

        /// <summary>
        /// 状况固定枚举
        /// </summary>
        public List<string> CatalogStatuses { get; } = [.. PlanValidation.ValidStatuses];

        /// <summary>
        /// 报告状态可选项（已完成 / 进行中 / 无要求）
        /// </summary>
        public List<string> ReportStatusOptions { get; } = [.. ReportStatusKind.All];

        private readonly DispatcherTimer _searchTimer;
        private string _searchKeyword;
        /// <summary>
        /// 搜索关键字（防抖）：一个搜索框同时过滤领退表与计划表
        /// </summary>
        public string SearchKeyword
        {
            get => _searchKeyword;
            set
            {
                if (SetProperty(ref _searchKeyword, value))
                {
                    _searchTimer.Stop();
                    _searchTimer.Start();
                }
            }
        }

        /* ###############################  列筛选（表头右键菜单）  ################################ */

        /// <summary>
        /// 取某列的筛选条件（没有则返回 null）
        /// </summary>
        public ColumnFilter GetColumnFilter(bool planTable, string property)
            => property != null && (planTable ? _planFilters : _reqFilters).TryGetValue(property, out ColumnFilter filter) ? filter : null;

        /// <summary>
        /// 该列是否有生效的筛选
        /// </summary>
        public bool IsColumnFiltered(bool planTable, string property) => GetColumnFilter(planTable, property)?.IsActive == true;

        /// <summary>
        /// 该表是否有任何生效的筛选
        /// </summary>
        public bool HasFilters(bool planTable) => (planTable ? _planFilters : _reqFilters).Values.Any(f => f.IsActive);

        /// <summary>
        /// 某列的去重取值（按表格显示文本，日期为 yyyy/M/d；数字/日期按数值排序）
        /// </summary>
        public List<string> GetColumnValues(bool planTable, string property)
        {
            IEnumerable<object> rows = planTable ? Plans.Cast<object>() : Requisitions.Cast<object>();
            List<string> values = rows.Select(r => FormatValue(r, property))
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Distinct(StringComparer.Ordinal)
                .ToList();
            values.Sort(CompareValues);
            return values;
        }

        /// <summary>
        /// 设置某列的筛选值（selected 为空或覆盖全部可选值时视为不过滤）
        /// </summary>
        public void SetColumnFilter(bool planTable, string property, string label, IReadOnlyList<string> options, IEnumerable<string> selected)
        {
            Dictionary<string, ColumnFilter> filters = planTable ? _planFilters : _reqFilters;
            HashSet<string> picked = new(selected ?? [], StringComparer.Ordinal);
            if (options != null && picked.Count >= options.Count)
            {
                // 全选 = 不筛选
                picked.Clear();
            }
            if (picked.Count == 0)
            {
                filters.Remove(property);
            }
            else
            {
                if (!filters.TryGetValue(property, out ColumnFilter filter))
                {
                    filter = new ColumnFilter(property, label);
                    filters[property] = filter;
                }
                filter.Label = label ?? filter.Label;
                filter.Options.Clear();
                if (options != null)
                {
                    filter.Options.AddRange(options);
                }
                filter.Selected.Clear();
                foreach (string value in picked)
                {
                    filter.Selected.Add(value);
                }
            }
            RefreshTable(planTable);
        }

        /// <summary>
        /// 清除某列筛选
        /// </summary>
        public void ClearColumnFilter(bool planTable, string property)
        {
            if ((planTable ? _planFilters : _reqFilters).Remove(property))
            {
                RefreshTable(planTable);
            }
        }

        /// <summary>
        /// 清除该表全部筛选
        /// </summary>
        public void ClearAllFilters(bool planTable)
        {
            Dictionary<string, ColumnFilter> filters = planTable ? _planFilters : _reqFilters;
            if (filters.Count == 0)
            {
                return;
            }
            filters.Clear();
            RefreshTable(planTable);
        }

        private void RefreshTable(bool planTable)
        {
            if (planTable)
            {
                PlansView.Refresh();
                UpdateStatusCounts();
            }
            else
            {
                RequisitionsView.Refresh();
            }
        }

        /// <summary>
        /// 行对象按列取值并格式化为表格显示文本
        /// </summary>
        private static string FormatValue(object row, string property)
        {
            if (row == null || string.IsNullOrWhiteSpace(property))
            {
                return "";
            }
            try
            {
                (Type, string) key = (row.GetType(), property);
                if (!PropertyCache.TryGetValue(key, out System.Reflection.PropertyInfo info))
                {
                    info = row.GetType().GetProperty(property);
                    PropertyCache[key] = info;
                }
                object value = info?.GetValue(row);
                return value switch
                {
                    null => "",
                    DateTime date => date.ToString("yyyy/M/d"),
                    _ => value.ToString()
                };
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// 取值排序：两边都能当数字按数值、都能当日期按日期，否则按当前区域字符串比较
        /// </summary>
        private static int CompareValues(string a, string b)
        {
            if (double.TryParse(a, out double na) && double.TryParse(b, out double nb))
            {
                return na.CompareTo(nb);
            }
            if (DateTime.TryParse(a, out DateTime da) && DateTime.TryParse(b, out DateTime db))
            {
                return da.CompareTo(db);
            }
            return string.Compare(a, b, StringComparison.CurrentCulture);
        }

        public PlansViewModel(DatabaseService db, PlanExcelService excelService, IPathService pathService,
            IPermissionService permission, ReviewService reviewService, AdminService adminService, AppSettingsService appSettings,
            ReportScanService reportScan, HighCostTaskCoordinator coordinator)
        {
            _db = db;
            _excelService = excelService;
            _pathService = pathService;
            _permission = permission;
            _reviewService = reviewService;
            _adminService = adminService;
            _appSettings = appSettings;
            _reportScan = reportScan;
            _coordinator = coordinator;

            PlansView = CollectionViewSource.GetDefaultView(Plans);
            PlansView.Filter = PlanFilter;
            PlansView.SortDescriptions.Add(new SortDescription(nameof(Plan.Id), ListSortDirection.Descending));
            RequisitionsView = CollectionViewSource.GetDefaultView(Requisitions);
            RequisitionsView.Filter = RequisitionFilter;
            RequisitionsView.SortDescriptions.Add(new SortDescription(nameof(Requisition.Id), ListSortDirection.Descending));

            _searchTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(250) };
            _searchTimer.Tick += (s, e) =>
            {
                _searchTimer.Stop();
                // 同一个搜索框：领退表与计划表一起刷新
                PlansView.Refresh();
                RequisitionsView.Refresh();
                UpdateStatusCounts();
            };

            // 报告扫描进度订阅：服务层事件已切回 UI 线程，这里直接刷属性
            _reportScan.Changed += OnReportScanChanged;
            _reportScan.ScanCompleted += OnReportScanCompletedInternal;

            Refresh();
        }

        /// <summary>
        /// 解除对单例 ReportScanService 事件的订阅（窗口关闭时调用）。
        /// PlansViewModel 是 Transient，ReportScanService 是 Singleton；不退订会造成
        /// 旧 ViewModel 被单例事件长期引用（泄露），且扫描回调会在多个旧实例上重复触发。
        /// </summary>
        public void DetachScanEvents()
        {
            _reportScan.Changed -= OnReportScanChanged;
            _reportScan.ScanCompleted -= OnReportScanCompletedInternal;
        }

        /* ###############################  排序（菜单驱动）  ################################ */

        /// <summary>
        /// 计划表排序。列头点击排序已关闭，排序统一由右键菜单/窗口菜单触发；
        /// propertyName 为空表示恢复默认顺序（按 Id 倒序，最新在前）。
        /// </summary>
        public void SortPlans(string propertyName, ListSortDirection direction)
        {
            ApplySort(PlansView, nameof(Plan.Id), propertyName, direction);
        }

        /// <summary>
        /// 领退表排序（说明同 <see cref="SortPlans"/>）
        /// </summary>
        public void SortRequisitions(string propertyName, ListSortDirection direction)
        {
            ApplySort(RequisitionsView, nameof(Requisition.Id), propertyName, direction);
        }

        private static void ApplySort(ICollectionView view, string defaultProperty, string propertyName, ListSortDirection direction)
        {
            view.SortDescriptions.Clear();
            view.SortDescriptions.Add(string.IsNullOrWhiteSpace(propertyName)
                ? new SortDescription(defaultProperty, ListSortDirection.Descending)
                : new SortDescription(propertyName, direction));
            view.Refresh();
        }

        /* ###############################  命令  ################################ */

        private RelayCommand _refreshCommand;
        public ICommand RefreshCommand => _refreshCommand ??= new RelayCommand(RefreshAndRescan);

        /// <summary>
        /// 刷新：重新加载数据（不再连带扫描报告文件夹——扫描已改为独立的高耗时任务，
        /// 由用户手动触发或闲置自动触发）
        /// </summary>
        private void RefreshAndRescan()
        {
            Refresh();
        }

        private RelayCommand _exportRequisitionCommand;
        public ICommand ExportRequisitionCommand => _exportRequisitionCommand ??= new RelayCommand(ExportRequisition);

        private RelayCommand _exportScheduleCommand;
        public ICommand ExportScheduleCommand => _exportScheduleCommand ??= new RelayCommand(ExportSchedule);

        private RelayCommand _clearAllCommand;
        /// <summary>
        /// 清空全部数据（已迁移至管理模块，仅管理员可操作）
        /// </summary>
        public ICommand ClearAllCommand => _clearAllCommand ??= new RelayCommand(ClearAll, () => false);

        private RelayCommand _saveChangesCommand;
        public ICommand SaveChangesCommand => _saveChangesCommand ??= new RelayCommand(SaveChanges, () => CanGridEdit && HasPendingChanges);

        private RelayCommand _discardChangesCommand;
        public ICommand DiscardChangesCommand => _discardChangesCommand ??= new RelayCommand(DiscardChanges, () => CanGridEdit && HasPendingChanges);

        private RelayCommand _addRequisitionCommand;
        /// <summary>
        /// 领退表新增（对话框，含计划表同步必填信息）
        /// </summary>
        public ICommand AddRequisitionCommand => _addRequisitionCommand ??= new RelayCommand(() => AddRequisition(), () => CanEdit);

        private RelayCommand _addPlanCommand;
        /// <summary>
        /// 计划表直接新增（QRT 前缀，非领用计划）
        /// </summary>
        public ICommand AddPlanCommand => _addPlanCommand ??= new RelayCommand(AddPlan, () => CanEdit);

        private RelayCommand _editRequisitionCommand;
        /// <summary>
        /// 领退表编辑
        /// </summary>
        public ICommand EditRequisitionCommand => _editRequisitionCommand ??= new RelayCommand(EditRequisition, () => CanEdit && SelectedRequisition != null);

        private RelayCommand _editPlanCommand;
        /// <summary>
        /// 计划表编辑
        /// </summary>
        public ICommand EditPlanCommand => _editPlanCommand ??= new RelayCommand(EditPlan, () => CanEdit && SelectedPlan != null);

        private CommunityToolkit.Mvvm.Input.RelayCommand<object> _deleteRequisitionCommand;
        /// <summary>
        /// 标记删除领退表行（参数为 Requisition）
        /// </summary>
        public ICommand DeleteRequisitionCommand => _deleteRequisitionCommand ??= new CommunityToolkit.Mvvm.Input.RelayCommand<object>(DeleteRequisition, p => CanGridEdit && p is Requisition);

        private CommunityToolkit.Mvvm.Input.RelayCommand<object> _deletePlanCommand;
        /// <summary>
        /// 标记删除计划表行（参数为 Plan）
        /// </summary>
        public ICommand DeletePlanCommand => _deletePlanCommand ??= new CommunityToolkit.Mvvm.Input.RelayCommand<object>(DeletePlan, p => CanGridEdit && p is Plan);

        private CommunityToolkit.Mvvm.Input.RelayCommand<object> _openSnFileCommand;
        /// <summary>
        /// 打开指定记录的SN文件（参数为 Requisition）
        /// </summary>
        public CommunityToolkit.Mvvm.Input.RelayCommand<object> OpenSnFileCommand
            => _openSnFileCommand ??= new CommunityToolkit.Mvvm.Input.RelayCommand<object>(OpenSnFile);

        /* ###############################  功能函数  ################################ */

        /// <summary>
        /// 从数据库重新加载两张表，并刷新筛选候选项（丢弃未提交修改）
        /// </summary>
        public void Refresh()
        {
            try
            {
                List<Plan> plans = _db.FreeSql.Select<Plan>().OrderByDescending(p => p.Id).ToList();
                Plans.Clear();
                _planOriginals.Clear();
                _pendingPlanAdded.Clear();
                _pendingPlanDeleted.Clear();
                foreach (Plan plan in plans)
                {
                    Plans.Add(plan);
                }

                List<Requisition> reqs = _db.FreeSql.Select<Requisition>().OrderByDescending(r => r.Id).ToList();
                Requisitions.Clear();
                _reqOriginals.Clear();
                _pendingReqAdded.Clear();
                _pendingReqDeleted.Clear();
                foreach (Requisition req in reqs)
                {
                    Requisitions.Add(req);
                    _reqOriginals[req.Id] = CloneReq(req);
                }

                LoadCatalogs();
                // 打开窗口仅从数据库加载报告夹扫描结果（不重新遍历文件系统，提速）；
                // 扫描改为独立的高耗时任务，由用户手动触发或闲置自动触发（见 ReportScanScheduler），
                // 不再在打开领退与计划窗口时同步执行
                LoadReportLinksFromDb();
                UpdatePlanReportFlags();
                // 计划表快照在报告标记（HasReportLink 已 JsonIgnore，不参与对比）加载后捕获
                foreach (Plan plan in Plans)
                {
                    _planOriginals[plan.Id] = ClonePlan(plan);
                }
                PlansView.Refresh();
                RequisitionsView.Refresh();
                UpdateStatusCounts();
                OnPropertyChanged(nameof(HasPendingChanges));
                OnPropertyChanged(nameof(PendingText));
                StatusMessage = string.Format(LanguageService.Get("Plans_StatusCount"), Requisitions.Count, Plans.Count);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "加载计划数据失败");
                StatusMessage = string.Format(LanguageService.Get("Plans_LoadFailed"), ex.Message);
            }
        }

        /// <summary>
        /// 加载编辑字典（测试项目/产品别/客户别/阶段）
        /// </summary>
        private void LoadCatalogs()
        {
            CatalogTestItems = _adminService.GetTestItems().Select(t => t.Name).ToList();
            CatalogProducts = _adminService.GetProducts();
            CatalogCustomers = _adminService.GetCustomers().Select(c => c.Name).ToList();
            CatalogStages = _adminService.GetStages().Select(s => s.Name).ToList();
            OnPropertyChanged(nameof(CatalogTestItems));
            OnPropertyChanged(nameof(CatalogProducts));
            OnPropertyChanged(nameof(CatalogCustomers));
            OnPropertyChanged(nameof(CatalogStages));
        }

        /* ###############################  报告扫描（按工作编号对应）  ################################ */

        /// <summary>
        /// 查找指定工作编号对应的报告夹信息（未找到返回 null）
        /// </summary>
        public ReportLink FindReportLink(string jobNo)
            => string.IsNullOrWhiteSpace(jobNo) || !_reportLinks.TryGetValue(jobNo, out ReportLink link) ? null : link;

        /// <summary>
        /// 从数据库加载已保存的报告夹扫描结果（打开窗口时调用，不遍历文件系统）
        /// </summary>
        private void LoadReportLinksFromDb()
        {
            _reportLinks.Clear();
            try
            {
                foreach (ReportLink link in _db.FreeSql.Select<ReportLink>().ToList())
                {
                    _reportLinks[link.JobNo] = link;
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"加载报告夹记录失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 查找计划记录对应的领退记录（右键打开一键报告时携带）。
        /// 匹配规则：备注含回线RT工令 → 备注含 WorkOrder → 机种相同且领用日期同开始日期。
        /// </summary>
        public Requisition FindRequisitionForPlan(Plan plan)
        {
            if (plan == null)
            {
                return null;
            }
            // 优先使用当前界面内存数据（含暂存新增/编辑），避免右键打开报告时读不到未提交的领退记录
            List<Requisition> reqs = Requisitions.ToList();
            return reqs.FirstOrDefault(r => !string.IsNullOrWhiteSpace(r.ReturnRtOrder)
                    && plan.Remark != null && plan.Remark.Contains(r.ReturnRtOrder))
                ?? reqs.FirstOrDefault(r => !string.IsNullOrWhiteSpace(r.WorkOrder)
                    && plan.Remark != null && plan.Remark.Contains(r.WorkOrder))
                ?? reqs.FirstOrDefault(r => r.ModelName == plan.ModelName
                    && r.RequisitionDate != null && plan.StartDate != null
                    && r.RequisitionDate.Value.Date == plan.StartDate.Value.Date);
        }

        /// <summary>
        /// 报告文件夹扫描完成（参数为匹配到的报告夹数量）。界面据此提示用户建立计划索引。
        /// </summary>
        public event Action<int> ReportScanCompleted;

        /// <summary>本次扫描的完成回调（手动更新报告文件夹时用来提示结果）</summary>
        private Action<int> _scanCallback;

        /* ###############################  报告扫描进度（状态栏绑定）  ################################ */

        private bool _scanProgressVisible;
        /// <summary>状态栏扫描进度区是否可见（正在扫描或被中断时可见）</summary>
        public bool ScanProgressVisible { get => _scanProgressVisible; set => SetProperty(ref _scanProgressVisible, value); }

        private string _scanProgressText;
        /// <summary>状态栏扫描进度文本（如「正在扫描报告文件夹 12/80：FSF050-9TAG …」）</summary>
        public string ScanProgressText { get => _scanProgressText; set => SetProperty(ref _scanProgressText, value); }

        private int _scanTotal;
        public int ScanTotal { get => _scanTotal; set => SetProperty(ref _scanTotal, value); }

        private int _scanProcessed;
        public int ScanProcessed { get => _scanProcessed; set => SetProperty(ref _scanProcessed, value); }

        private RelayCommand _stopScanCommand;
        /// <summary>停止扫描（发取消信号，当前报告夹读完再停）</summary>
        public ICommand StopScanCommand => _stopScanCommand ??= new RelayCommand(
            () => _coordinator.RequestInterruptCurrent(),
            () => _reportScan.IsRunning);

        /// <summary>
        /// 手动触发报告扫描（走协调器，与计划索引/一键报告互斥）。
        /// </summary>
        /// <param name="completed">扫描完成后的回调（参数为匹配到的报告夹数量）</param>
        public void RequestReportScan(Action<int> completed = null)
        {
            _scanCallback = completed;
            _ = _coordinator.RequestStartAsync(LanguageService.Get("ReportScan_TaskName"),
                cts => _reportScan.RunAsync(cts.Token));
        }

        /// <summary>
        /// 扫描进度变化（服务层已切回 UI 线程）：刷新状态栏绑定属性
        /// </summary>
        private void OnReportScanChanged()
        {
            ScanProgressVisible = _reportScan.IsRunning || _reportScan.WasInterrupted;
            ScanTotal = Math.Max(_reportScan.Total, 1);
            ScanProcessed = Math.Min(_reportScan.Processed, ScanTotal);
            ScanProgressText = _reportScan.IsRunning
                ? string.Format(LanguageService.Get("ReportScan_ProgressFormat"),
                    _reportScan.Processed, _reportScan.Total, _reportScan.CurrentFolder ?? "")
                : (_reportScan.WasInterrupted
                    ? LanguageService.Get("ReportScan_Interrupted")
                    : "");
            _stopScanCommand?.RaiseCanExecuteChanged();
            // 被中断后稍作停留再隐藏，让用户看到「已中断」反馈
            if (!_reportScan.IsRunning && _reportScan.WasInterrupted)
            {
                DispatcherTimer hideTimer = new() { Interval = TimeSpan.FromSeconds(3) };
                hideTimer.Tick += (s, e) =>
                {
                    hideTimer.Stop();
                    if (!_reportScan.IsRunning)
                    {
                        ScanProgressVisible = false;
                    }
                };
                hideTimer.Start();
            }
        }

        /// <summary>
        /// 扫描完成（服务层已切回 UI 线程）：重载 report_links 缓存与计划表报告标记，
        /// 并重新拉取 plans（ReportStatus 可能被扫描改写）
        /// </summary>
        private void OnReportScanCompletedInternal(int matchedCount)
        {
            try
            {
                LoadReportLinksFromDb();
                // ReportStatus 被扫描写回了数据库，重新拉取 plans 以刷新表格显示
                List<Plan> plans = _db.FreeSql.Select<Plan>().OrderByDescending(p => p.Id).ToList();
                Plans.Clear();
                foreach (Plan plan in plans)
                {
                    Plans.Add(plan);
                }
                UpdatePlanReportFlags();
                // 计划表快照同步更新（扫描写回的 ReportStatus 不算用户未提交修改）
                _planOriginals.Clear();
                foreach (Plan plan in Plans)
                {
                    _planOriginals[plan.Id] = ClonePlan(plan);
                }
                PlansView.Refresh();
                UpdateStatusCounts();
                OnPropertyChanged(nameof(HasPendingChanges));
                OnPropertyChanged(nameof(PendingText));
            }
            catch (Exception ex)
            {
                _logger.Warn($"扫描完成后刷新计划表失败: {ex.Message}");
            }
            // 转发给界面（提示建立计划索引等）与手动触发的回调
            ReportScanCompleted?.Invoke(matchedCount);
            Action<int> callback = _scanCallback;
            _scanCallback = null;
            callback?.Invoke(matchedCount);
        }

        /// <summary>
        /// 按计划表工作编号是否有对应报告夹刷新显示标记（无对应时界面用近黑色区分）
        /// </summary>
        private void UpdatePlanReportFlags()
        {
            foreach (Plan plan in Plans)
            {
                plan.HasReportLink = !string.IsNullOrWhiteSpace(plan.JobNo) && _reportLinks.ContainsKey(plan.JobNo);
            }
        }

        private static Plan ClonePlan(Plan source)
            => JsonConvert.DeserializeObject<Plan>(JsonConvert.SerializeObject(source));

        private static Requisition CloneReq(Requisition source)
            => JsonConvert.DeserializeObject<Requisition>(JsonConvert.SerializeObject(source));

        /// <summary>
        /// 列筛选是否通过（值以表格显示文本比较；Selected 为空表示该列不过滤）
        /// </summary>
        private static bool PassesColumnFilters(IEnumerable<ColumnFilter> filters, object row)
        {
            foreach (ColumnFilter filter in filters)
            {
                if (filter.IsActive && !filter.Selected.Contains(FormatValue(row, filter.Property)))
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// 计划表筛选：本表列筛选 + 公共搜索关键字
        /// </summary>
        private bool PlanFilter(object obj)
        {
            if (obj is not Plan plan || !PassesColumnFilters(_planFilters.Values, plan))
            {
                return false;
            }
            if (string.IsNullOrWhiteSpace(SearchKeyword))
            {
                return true;
            }
            string kw = SearchKeyword.Trim();
            return Contains(plan.ModelName, kw) || Contains(plan.JobNo, kw)
                || Contains(plan.TestItem, kw) || Contains(plan.Owner, kw)
                || Contains(plan.Product, kw) || Contains(plan.Customer, kw);
        }

        /// <summary>
        /// 领退表筛选：本表列筛选（与计划表独立）+ 公共搜索关键字
        /// </summary>
        private bool RequisitionFilter(object obj)
        {
            if (obj is not Requisition req || !PassesColumnFilters(_reqFilters.Values, req))
            {
                return false;
            }
            if (string.IsNullOrWhiteSpace(SearchKeyword))
            {
                return true;
            }
            string kw = SearchKeyword.Trim();
            return Contains(req.ModelName, kw) || Contains(req.RequisitionNo, kw)
                || Contains(req.WorkOrder, kw) || Contains(req.ReturnRtOrder, kw)
                || Contains(req.SN, kw) || Contains(req.StockInNo, kw);
        }

        private static bool Contains(string source, string keyword)
            => source?.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0;

        /* ###############################  单元格编辑  ################################ */

        /// <summary>
        /// 单元格校验（工作编号格式/状况枚举/字典存在性）；合法返回null
        /// </summary>
        public string ValidateField(string field, string value)
        {
            return field switch
            {
                "JobNo" => PlanValidation.ValidateJobNo(value),
                "Status" => PlanValidation.ValidateStatus(value),
                "TestItem" => PlanValidation.ValidateInCatalog(value, CatalogTestItems, "测试项目"),
                "Product" => PlanValidation.ValidateInCatalog(value, CatalogProducts, "产品别"),
                "Customer" => PlanValidation.ValidateInCatalog(value, CatalogCustomers, "客户别"),
                "Stage" => PlanValidation.ValidateInCatalog(value, CatalogStages, "阶段"),
                _ => null
            };
        }

        /// <summary>
        /// 机种联动：输入机种名称后自动带出产品别/客户别（仅填充空字段）。
        /// 查询规则：产品别 = 机种名开始 2 位代码，客户别 = 机种名第 8 位起的 2 位代码；代码映射缺失时回退机种映射表。
        /// </summary>
        public void AutoFillByModel(Plan plan)
        {
            if (string.IsNullOrWhiteSpace(plan.ModelName))
            {
                return;
            }
            string product = _adminService.FindProductByModel(plan.ModelName);
            string customer = _adminService.FindCustomerByModel(plan.ModelName);
            if (product == null || customer == null)
            {
                ModelMapping mapping = _adminService.FindModelMapping(plan.ModelName);
                product ??= mapping?.Product;
                customer ??= mapping?.Customer;
            }
            if (string.IsNullOrWhiteSpace(plan.Product) && product != null)
            {
                plan.Product = product;
            }
            if (string.IsNullOrWhiteSpace(plan.Customer) && customer != null)
            {
                plan.Customer = customer;
            }
        }

        /// <summary>
        /// 测试项目联动：选择测试项目后自动带出负责人/试验时间，并按开始日期+试验时间计算结束日期
        /// </summary>
        public void AutoFillByTestItem(Plan plan)
        {
            if (string.IsNullOrWhiteSpace(plan.TestItem))
            {
                return;
            }
            TestItemCatalog item = _adminService.GetTestItems().FirstOrDefault(t => t.Name == plan.TestItem);
            if (item == null)
            {
                return;
            }
            if (!string.IsNullOrWhiteSpace(item.Owner))
            {
                plan.Owner = item.Owner;
            }
            if (!string.IsNullOrWhiteSpace(item.Period))
            {
                plan.TestPeriod = item.Period;
            }
            if (plan.StartDate != null && int.TryParse(item.Period, out int hours))
            {
                plan.EndDate = plan.StartDate.Value.AddHours(hours);
            }
        }

        /// <summary>
        /// 标记单元格已修改（用于刷新待提交状态提示）
        /// </summary>
        public void NotifyPendingChanged()
        {
            OnPropertyChanged(nameof(HasPendingChanges));
            OnPropertyChanged(nameof(PendingText));
            UpdateStatusCounts();
            CommandManager.InvalidateRequerySuggested();
        }

        /* ###############################  增删改  ################################ */

        /// <summary>
        /// 领退表新增（对话框，含计划表同步必填信息）。
        /// prefillPlan 非空时用于「计划表新增 → 转为领用」：把计划表那侧已填的内容带进领退表窗口。
        /// </summary>
        private void AddRequisition(Plan prefillPlan = null)
        {
            if (!CanEdit)
            {
                StatusMessage = LanguageService.Get("Plans_NoAddPermission");
                return;
            }
            Views.WindowRequisitionEdit editWindow = new(_db, _permission, _adminService, _excelService, null);
            if (prefillPlan != null)
            {
                editWindow.PrefillFromPlan(prefillPlan);
            }
            // 非模态：保存后通过 Saved 事件回调处理暂存/提审，窗口打开期间主界面仍可操作
            editWindow.Saved += (reqResult, planResult, editId) =>
            {
                if (NeedsReview)
                {
                    _reviewService.SubmitPlanRequest("新增", planResult, null, _permission.CurrentUser);
                    _reviewService.SubmitRequisitionRequest("新增", reqResult, null, _permission.CurrentUser);
                    StatusMessage = LanguageService.Get("Plans_AddSubmitted");
                    _ = System.Windows.MessageBox.Show(LocalizationHelper.Get("Msg_AddSubmitted"), LanguageService.Get("Cap_SubmitSuccess"));
                }
                else
                {
                    // 暂存到内存（需点“提交保存”才写库）
                    _pendingReqAdded.Add(reqResult);
                    _pendingPlanAdded.Add(planResult);
                    Requisitions.Insert(0, reqResult);
                    Plans.Insert(0, planResult);
                    NotifyPendingChanged();
                    StatusMessage = PendingText;
                }
            };
            editWindow.Show();
        }

        private void AddPlan()
        {
            if (!CanEdit)
            {
                StatusMessage = LanguageService.Get("Plans_NoAddPermission");
                return;
            }
            Views.WindowPlanDirectEdit editWindow = new(_db, _permission, _adminService, _excelService, null);
            // 「转为领用」：把计划表窗口里填好的内容带进领退表新增窗口，
            // 保存后由领用流程建立 RT 计划（当前这个 QRT 计划不入库，直接关掉）
            editWindow.ConvertToRequisitionRequested += planDraft => AddRequisition(planDraft);
            editWindow.Saved += (planResult, editId) =>
            {
                if (NeedsReview)
                {
                    _reviewService.SubmitPlanRequest("新增", planResult, null, _permission.CurrentUser);
                    StatusMessage = LanguageService.Get("Plans_AddSubmitted");
                    _ = System.Windows.MessageBox.Show(LocalizationHelper.Get("Msg_AddSubmitted"), LanguageService.Get("Cap_SubmitSuccess"));
                }
                else
                {
                    // 暂存到内存（需点“提交保存”才写库）
                    _pendingPlanAdded.Add(planResult);
                    Plans.Insert(0, planResult);
                    NotifyPendingChanged();
                    StatusMessage = PendingText;
                }
            };
            editWindow.Show();
        }

        private void EditRequisition()
        {
            if (!CanEdit || SelectedRequisition == null)
            {
                return;
            }
            Requisition target = SelectedRequisition;
            Views.WindowRequisitionEdit editWindow = new(_db, _permission, _adminService, _excelService, target);
            editWindow.Saved += (reqResult, planResult, editId) =>
            {
                if (NeedsReview)
                {
                    _reviewService.SubmitRequisitionRequest("编辑", reqResult, target.Id, _permission.CurrentUser);
                    if (planResult != null)
                    {
                        _reviewService.SubmitPlanRequest("编辑", planResult, planResult.Id, _permission.CurrentUser);
                    }
                    StatusMessage = LanguageService.Get("Plans_EditSubmitted");
                    _ = System.Windows.MessageBox.Show(LocalizationHelper.Get("Msg_EditSubmitted"), LanguageService.Get("Cap_SubmitSuccess"));
                }
                else
                {
                    // 暂存到内存：把对话框结果复制回集合中的对象（快照对比将识别为修改）
                    CopyRequisitionFields(reqResult, target);
                    if (planResult != null)
                    {
                        // 同步暂存关联计划的修改（按 Id 找到集合内对象）
                        Plan existingPlan = Plans.FirstOrDefault(p => p.Id == planResult.Id);
                        if (existingPlan != null)
                        {
                            CopyPlanFields(planResult, existingPlan);
                        }
                    }
                    NotifyPendingChanged();
                    StatusMessage = PendingText;
                }
            };
            editWindow.Show();
        }

        private void EditPlan()
        {
            if (!CanEdit || SelectedPlan == null)
            {
                return;
            }
            Plan target = SelectedPlan;
            Views.WindowPlanDirectEdit editWindow = new(_db, _permission, _adminService, _excelService, target);
            editWindow.Saved += (planResult, editId) =>
            {
                if (NeedsReview)
                {
                    _reviewService.SubmitPlanRequest("编辑", planResult, target.Id, _permission.CurrentUser);
                    StatusMessage = LanguageService.Get("Plans_EditSubmitted");
                    _ = System.Windows.MessageBox.Show(LocalizationHelper.Get("Msg_EditSubmitted"), LanguageService.Get("Cap_SubmitSuccess"));
                }
                else
                {
                    // 暂存到内存：把对话框结果复制回集合中的对象（快照对比将识别为修改）
                    CopyPlanFields(planResult, target);
                    NotifyPendingChanged();
                    StatusMessage = PendingText;
                }
            };
            editWindow.Show();
        }

        /// <summary>
        /// 复制领退记录字段（保持集合内对象引用，触发属性通知自动刷新单元格）
        /// </summary>
        private static void CopyRequisitionFields(Requisition from, Requisition to)
        {
            to.RequisitionDate = from.RequisitionDate;
            to.RequisitionNo = from.RequisitionNo;
            to.ModelName = from.ModelName;
            to.OutQty = from.OutQty;
            to.SN = from.SN;
            to.SnFilePath = from.SnFilePath;
            to.Rev = from.Rev;
            to.WorkOrder = from.WorkOrder;
            to.DC = from.DC;
            to.LineNo = from.LineNo;
            to.ReturnRtOrder = from.ReturnRtOrder;
            to.ReturnQty = from.ReturnQty;
            to.ReturnDate = from.ReturnDate;
            to.StockInNo = from.StockInNo;
            to.StockInQty = from.StockInQty;
            to.StockInDate = from.StockInDate;
            to.Remark = from.Remark;
            to.UpdatedBy = from.UpdatedBy;
            to.UpdatedAt = from.UpdatedAt;
        }

        /// <summary>
        /// 复制计划记录字段（保持集合内对象引用，触发属性通知自动刷新单元格）
        /// </summary>
        private static void CopyPlanFields(Plan from, Plan to)
        {
            to.JobNo = from.JobNo;
            to.Product = from.Product;
            to.Customer = from.Customer;
            to.ModelName = from.ModelName;
            to.Stage = from.Stage;
            to.TestItem = from.TestItem;
            to.SampleSize = from.SampleSize;
            to.TestPeriod = from.TestPeriod;
            to.Owner = from.Owner;
            to.StartDate = from.StartDate;
            to.EndDate = from.EndDate;
            to.Status = from.Status;
            to.ReportStatus = from.ReportStatus;
            to.Remark = from.Remark;
            to.UpdatedBy = from.UpdatedBy;
            to.UpdatedAt = from.UpdatedAt;
        }

        private void DeleteRequisition(object parameter)
        {
            if (parameter is not Requisition req || !CanGridEdit)
            {
                return;
            }
            string reqNo = req.RequisitionNo ?? "-";
            string reqModel = req.ModelName ?? "-";
            if (System.Windows.MessageBox.Show(
                $"确认删除该领退记录？（暂存，提交后生效）\n領料單据號: {reqNo}  机种: {reqModel}", LanguageService.Get("Cap_DeleteConfirm"), System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Warning)
                != System.Windows.MessageBoxResult.Yes)
            {
                return;
            }
            if (req.Id == 0)
            {
                _pendingReqAdded.Remove(req);
                Requisitions.Remove(req);
            }
            else
            {
                _pendingReqDeleted[req.Id] = req;
                Requisitions.Remove(req);
            }
            NotifyPendingChanged();
            StatusMessage = PendingText;
        }

        private void DeletePlan(object parameter)
        {
            if (parameter is not Plan plan || !CanGridEdit)
            {
                return;
            }
            string jobNo = plan.JobNo ?? "-";
            string planModel = plan.ModelName ?? "-";
            if (System.Windows.MessageBox.Show(
                $"确认删除该计划记录？（暂存，提交后生效）\n工作編號: {jobNo}  机种: {planModel}", LanguageService.Get("Cap_DeleteConfirm"), System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Warning)
                != System.Windows.MessageBoxResult.Yes)
            {
                return;
            }
            if (plan.Id == 0)
            {
                _pendingPlanAdded.Remove(plan);
                Plans.Remove(plan);
            }
            else
            {
                _pendingPlanDeleted[plan.Id] = plan;
                Plans.Remove(plan);
            }
            NotifyPendingChanged();
            StatusMessage = PendingText;
        }

        private void OpenSnFile(object parameter)
        {
            if (parameter is not Requisition req || string.IsNullOrWhiteSpace(req.SnFilePath))
            {
                return;
            }
            string path = _db.ResolveAttachmentPath(req.SnFilePath);
            if (!File.Exists(path))
            {
                StatusMessage = string.Format(LanguageService.Get("Plans_SNFileNotExist"), path);
                _ = System.Windows.MessageBox.Show(string.Format(LanguageService.Get("Plans_SNFileNotExistMsg"), path), LanguageService.Get("Cap_Info"));
                return;
            }
            try
            {
                System.Diagnostics.Process.Start(path);
                _logger.Info($"打开SN文件: {path}");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"打开SN文件失败: {path}");
                StatusMessage = $"打开失败: {ex.Message}";
            }
        }

        /* ###############################  提交与丢弃  ################################ */

        /// <summary>
        /// 以快照对比检测计划表已修改的存量行数
        /// </summary>
        private int DetectPlanModifiedCount()
        {
            int count = 0;
            foreach (Plan plan in Plans.Where(p => p.Id > 0))
            {
                if (_planOriginals.TryGetValue(plan.Id, out Plan before)
                    && JsonConvert.SerializeObject(plan) != JsonConvert.SerializeObject(before))
                {
                    count++;
                }
            }
            return count;
        }

        /// <summary>
        /// 以快照对比检测领退表已修改的存量行数
        /// </summary>
        private int DetectReqModifiedCount()
        {
            int count = 0;
            foreach (Requisition req in Requisitions.Where(r => r.Id > 0))
            {
                if (_reqOriginals.TryGetValue(req.Id, out Requisition before)
                    && JsonConvert.SerializeObject(req) != JsonConvert.SerializeObject(before))
                {
                    count++;
                }
            }
            return count;
        }

        /// <summary>
        /// 丢弃所有暂存修改，还原为数据库状态
        /// </summary>
        private void DiscardChanges()
        {
            if (System.Windows.MessageBox.Show(LocalizationHelper.Get("Msg_ConfirmDiscard"), LanguageService.Get("Cap_DiscardConfirm"), System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Question)
                != System.Windows.MessageBoxResult.Yes)
            {
                return;
            }
            Refresh();
            StatusMessage = "已丢弃未提交的修改";
        }

        /// <summary>
        /// 提交保存：将暂存的新增/修改/删除写入数据库，并为每条变更写入变更日志
        /// </summary>
        private void SaveChanges()
        {
            if (!CanGridEdit || !HasPendingChanges)
            {
                return;
            }
            try
            {
                string op = _permission.CurrentUser;
                int added = 0, modified = 0, deleted = 0;

                foreach (Plan plan in _pendingPlanAdded)
                {
                    if (string.IsNullOrWhiteSpace(plan.JobNo))
                    {
                        StatusMessage = "存在未填写工作編號的计划空行，请补充或删除后再提交";
                        return;
                    }
                    plan.Id = _db.FreeSql.Insert(plan).ExecuteIdentity();
                    WritePlanLog("新增", plan.Id, $"新增计划 {plan.JobNo} ({plan.ModelName})", null, plan, op);
                    added++;
                }
                foreach (Plan plan in Plans.Where(p => p.Id > 0))
                {
                    if (!_planOriginals.TryGetValue(plan.Id, out Plan before)) continue;
                    if (JsonConvert.SerializeObject(plan) == JsonConvert.SerializeObject(before)) continue;
                    plan.UpdatedBy = op;
                    plan.UpdatedAt = DateTime.Now;
                    _db.FreeSql.Update<Plan>().SetSource(plan).Where(p => p.Id == plan.Id).ExecuteAffrows();
                    WritePlanLog("编辑", plan.Id, $"编辑计划 {plan.JobNo} ({plan.ModelName})", before, plan, op);
                    modified++;
                }
                foreach (KeyValuePair<long, Plan> kv in _pendingPlanDeleted)
                {
                    _db.FreeSql.Delete<Plan>().Where(p => p.Id == kv.Key).ExecuteAffrows();
                    WritePlanLog("删除", kv.Key, $"删除计划 {kv.Value.JobNo} ({kv.Value.ModelName})", kv.Value, null, op);
                    deleted++;
                }

                foreach (Requisition req in _pendingReqAdded)
                {
                    if (string.IsNullOrWhiteSpace(req.RequisitionNo))
                    {
                        StatusMessage = "存在未填写領料單据號的领退空行，请补充或删除后再提交";
                        return;
                    }
                    req.Id = _db.FreeSql.Insert(req).ExecuteIdentity();
                    WriteReqLog("新增", req.Id, $"新增领退 {req.RequisitionNo} ({req.ModelName})", null, req, op);
                    added++;
                }
                foreach (Requisition req in Requisitions.Where(r => r.Id > 0))
                {
                    if (!_reqOriginals.TryGetValue(req.Id, out Requisition before)) continue;
                    if (JsonConvert.SerializeObject(req) == JsonConvert.SerializeObject(before)) continue;
                    req.UpdatedBy = op;
                    req.UpdatedAt = DateTime.Now;
                    _db.FreeSql.Update<Requisition>().SetSource(req).Where(r => r.Id == req.Id).ExecuteAffrows();
                    WriteReqLog("编辑", req.Id, $"编辑领退 {req.RequisitionNo} ({req.ModelName})", before, req, op);
                    modified++;
                }
                foreach (KeyValuePair<long, Requisition> kv in _pendingReqDeleted)
                {
                    _db.FreeSql.Delete<Requisition>().Where(r => r.Id == kv.Key).ExecuteAffrows();
                    WriteReqLog("删除", kv.Key, $"删除领退 {kv.Value.RequisitionNo} ({kv.Value.ModelName})", kv.Value, null, op);
                    deleted++;
                }

                _logger.Info($"提交保存: 新增{added} 修改{modified} 删除{deleted} 操作人={op}");
                Refresh();
                StatusMessage = $"保存成功: 新增{added} 修改{modified} 删除{deleted}";
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "提交保存失败");
                StatusMessage = $"保存失败: {ex.Message}";
                _ = System.Windows.MessageBox.Show($"保存失败:\n{ex.Message}", LanguageService.Get("Cap_Error"));
            }
        }

        private void WritePlanLog(string action, long planId, string summary, Plan before, Plan after, string op)
        {
            _db.FreeSql.Insert(new PlanChangeLog
            {
                Action = action,
                PlanId = planId,
                Summary = summary,
                BeforeJson = before == null ? null : JsonConvert.SerializeObject(before),
                AfterJson = after == null ? null : JsonConvert.SerializeObject(after),
                Operator = op,
                CreatedAt = DateTime.Now
            }).ExecuteAffrows();
        }

        private void WriteReqLog(string action, long planId, string summary, Requisition before, Requisition after, string op)
        {
            _db.FreeSql.Insert(new PlanChangeLog
            {
                Action = action,
                PlanId = planId,
                Summary = summary,
                BeforeJson = before == null ? null : JsonConvert.SerializeObject(before),
                AfterJson = after == null ? null : JsonConvert.SerializeObject(after),
                Operator = op,
                CreatedAt = DateTime.Now
            }).ExecuteAffrows();
        }

        private void ExportRequisition()
        {
            if (!_permission.Can("plan.export"))
            {
                StatusMessage = "无导出权限";
                return;
            }
            string file = _pathService.SavePathDialog("导出领退表", $"{DateTime.Now:yyyy}.成品領用記錄.xlsx");
            if (file == null)
            {
                return;
            }
            try
            {
                _excelService.ExportRequisition(file);
                OleEmbedResult embed = _excelService.LastEmbedResult;
                StatusMessage = $"领退表已导出: {file}";
                // 附件（OLE）是否真的写进文件要如实告知：Excel 环境异常时可能只是把数据导出成功
                if (embed != null && !embed.Saved)
                {
                    StatusMessage += $"（{embed.Summary}）";
                    _ = System.Windows.MessageBox.Show($"{embed.Summary}\n文件已导出：{file}\n详情见日志。",
                        LanguageService.Get("Cap_Warning"), System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "导出领退表失败");
                _ = System.Windows.MessageBox.Show($"导出领退表失败:\n{ex.Message}", LanguageService.Get("Cap_Error"));
            }
        }

        private void ExportSchedule()
        {
            if (!_permission.Can("plan.export"))
            {
                StatusMessage = "无导出权限";
                return;
            }
            string file = _pathService.SavePathDialog("导出计划表", $"Y{DateTime.Now:yyyy} ORT Test Schedule.xlsx");
            if (file == null)
            {
                return;
            }
            try
            {
                _excelService.ExportSchedule(file);
                StatusMessage = $"计划表已导出: {file}";
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "导出计划表失败");
                _ = System.Windows.MessageBox.Show($"导出计划表失败:\n{ex.Message}", LanguageService.Get("Cap_Error"));
            }
        }

        private void ClearAll()
        {
            if (!_permission.Can("plan.delete"))
            {
                StatusMessage = "无删除权限";
                return;
            }
            if (Plans.Count == 0 && Requisitions.Count == 0)
            {
                StatusMessage = "没有可清空的数据";
                return;
            }
            if (System.Windows.MessageBox.Show($"确认清空全部计划数据（计划表{Plans.Count}条/领退表{Requisitions.Count}条）？此操作不可恢复！", LanguageService.Get("Cap_ClearConfirm"), System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Warning)
                != System.Windows.MessageBoxResult.Yes)
            {
                return;
            }
            try
            {
                int n = _excelService.ClearAll();
                Refresh();
                StatusMessage = $"已清空 {n} 条记录";
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "清空数据失败");
                StatusMessage = $"清空失败: {ex.Message}";
            }
        }
    }
}
