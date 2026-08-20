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
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Threading;

namespace ORT一键报告.Plans.ViewModels
{
    /// <summary>
    /// 单个筛选条件：字段 + 选中值集合（值字段）或日期范围（日期字段）。
    /// </summary>
    public class FilterCondition : ObservableObject
    {
        private readonly Action _onChanged;

        /// <summary>
        /// 字段名
        /// </summary>
        public string Field { get; }

        /// <summary>
        /// 是否为日期字段
        /// </summary>
        public bool IsDateField { get; }

        /// <summary>
        /// 非日期字段取反，便于XAML绑定Visibility
        /// </summary>
        public bool IsValueField => !IsDateField;

        private List<string> _options = [];
        /// <summary>
        /// 值字段的候选项（随数据动态更新）
        /// </summary>
        public List<string> Options { get => _options; set => SetProperty(ref _options, value); }

        /// <summary>
        /// 值字段的选中集合（空集合表示"全部"）
        /// </summary>
        public ObservableCollection<string> SelectedValues { get; } = [];

        private DateTime? _dateFrom;
        /// <summary>
        /// 日期范围起
        /// </summary>
        public DateTime? DateFrom
        {
            get => _dateFrom;
            set { if (SetProperty(ref _dateFrom, value)) _onChanged?.Invoke(); }
        }

        private DateTime? _dateTo;
        /// <summary>
        /// 日期范围止
        /// </summary>
        public DateTime? DateTo
        {
            get => _dateTo;
            set { if (SetProperty(ref _dateTo, value)) _onChanged?.Invoke(); }
        }

        public FilterCondition(string field, bool isDateField, Action onChanged)
        {
            Field = field;
            IsDateField = isDateField;
            _onChanged = onChanged;
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

        /// <summary>
        /// 计划表可筛选的字段定义：(字段名, 是否日期字段)
        /// </summary>
        private static readonly (string Name, bool IsDate)[] FilterFields =
        [
            ("机种", false), ("测试项目", false), ("产品别", false),
            ("客户别", false), ("负责人", false), ("阶段", false), ("状况", false),
            ("开始日期", true)
        ];

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
        /// 当前激活的筛选条件（计划表）
        /// </summary>
        public ObservableCollection<FilterCondition> ActiveFilters { get; } = [];

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
            ? $"未提交修改: 领退新增{_pendingReqAdded.Count} 计划新增{_pendingPlanAdded.Count} 修改{DetectPlanModifiedCount() + DetectReqModifiedCount()} 删除{_pendingPlanDeleted.Count + _pendingReqDeleted.Count}"
            : "无未提交修改";

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

        private readonly DispatcherTimer _searchTimer;
        private string _keyword;
        /// <summary>
        /// 计划表搜索关键字（防抖）
        /// </summary>
        public string Keyword
        {
            get => _keyword;
            set
            {
                if (SetProperty(ref _keyword, value))
                {
                    _searchTimer.Stop();
                    _searchTimer.Start();
                }
            }
        }

        private string _reqKeyword;
        /// <summary>
        /// 领退表搜索关键字
        /// </summary>
        public string ReqKeyword
        {
            get => _reqKeyword;
            set
            {
                if (SetProperty(ref _reqKeyword, value))
                {
                    RequisitionsView.Refresh();
                }
            }
        }

        private string _addFilterField;
        /// <summary>
        /// "添加筛选条件"下拉的选择项；选中后创建对应条件并复位
        /// </summary>
        public string AddFilterField
        {
            get => _addFilterField;
            set
            {
                _addFilterField = null;
                OnPropertyChanged(nameof(AddFilterField));
                if (value != null)
                {
                    AddFilter(value);
                }
            }
        }

        /// <summary>
        /// 尚未添加的可选筛选字段
        /// </summary>
        public List<string> AvailableFields
            => FilterFields.Select(f => f.Name).Where(n => ActiveFilters.All(c => c.Field != n)).ToList();

        public PlansViewModel(DatabaseService db, PlanExcelService excelService, IPathService pathService,
            IPermissionService permission, ReviewService reviewService, AdminService adminService)
        {
            _db = db;
            _excelService = excelService;
            _pathService = pathService;
            _permission = permission;
            _reviewService = reviewService;
            _adminService = adminService;

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
                PlansView.Refresh();
            };

            Refresh();
        }

        /* ###############################  命令  ################################ */

        private RelayCommand _refreshCommand;
        public ICommand RefreshCommand => _refreshCommand ??= new RelayCommand(Refresh);

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
        public ICommand AddRequisitionCommand => _addRequisitionCommand ??= new RelayCommand(AddRequisition, () => CanEdit);

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

        private CommunityToolkit.Mvvm.Input.RelayCommand<object> _removeFilterCommand;
        /// <summary>
        /// 删除指定筛选条件（参数为 FilterCondition）
        /// </summary>
        public ICommand RemoveFilterCommand => _removeFilterCommand ??= new CommunityToolkit.Mvvm.Input.RelayCommand<object>(RemoveFilter);

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
                    _planOriginals[plan.Id] = ClonePlan(plan);
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

                foreach (FilterCondition cond in ActiveFilters.Where(c => !c.IsDateField))
                {
                    RefreshConditionOptions(cond);
                }
                LoadCatalogs();
                PlansView.Refresh();
                RequisitionsView.Refresh();
                OnPropertyChanged(nameof(HasPendingChanges));
                OnPropertyChanged(nameof(PendingText));
                StatusMessage = $"领退表 {Requisitions.Count} 条 / 计划表 {Plans.Count} 条";
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "加载计划数据失败");
                StatusMessage = $"加载失败: {ex.Message}";
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

        private static Plan ClonePlan(Plan source)
            => JsonConvert.DeserializeObject<Plan>(JsonConvert.SerializeObject(source));

        private static Requisition CloneReq(Requisition source)
            => JsonConvert.DeserializeObject<Requisition>(JsonConvert.SerializeObject(source));

        private void AddFilter(string fieldName)
        {
            (string Name, bool IsDate) field = FilterFields.FirstOrDefault(f => f.Name == fieldName);
            if (field.Name == null)
            {
                return;
            }
            FilterCondition cond = new(field.Name, field.IsDate, () => PlansView.Refresh());
            if (!field.IsDate)
            {
                RefreshConditionOptions(cond);
                cond.SelectedValues.CollectionChanged += (s, e) => PlansView.Refresh();
            }
            ActiveFilters.Add(cond);
            OnPropertyChanged(nameof(AvailableFields));
            PlansView.Refresh();
        }

        private void RemoveFilter(object parameter)
        {
            if (parameter is FilterCondition cond && ActiveFilters.Remove(cond))
            {
                OnPropertyChanged(nameof(AvailableFields));
                PlansView.Refresh();
            }
        }

        private void RefreshConditionOptions(FilterCondition cond)
        {
            List<string> options = DistinctOptions(Plans.Select(GetFieldValue(cond.Field)));
            cond.Options = options;
            for (int i = cond.SelectedValues.Count - 1; i >= 0; i--)
            {
                if (!options.Contains(cond.SelectedValues[i]))
                {
                    cond.SelectedValues.RemoveAt(i);
                }
            }
        }

        private static List<string> DistinctOptions(IEnumerable<string> values)
        {
            List<string> options = [];
            options.AddRange(values.Where(v => !string.IsNullOrWhiteSpace(v)).Distinct().OrderBy(v => v));
            return options;
        }

        /// <summary>
        /// 字段名 -> 计划表取值函数
        /// </summary>
        private static Func<Plan, string> GetFieldValue(string field) => field switch
        {
            "机种" => p => p.ModelName,
            "测试项目" => p => p.TestItem,
            "产品别" => p => p.Product,
            "客户别" => p => p.Customer,
            "负责人" => p => p.Owner,
            "阶段" => p => p.Stage,
            "状况" => p => p.Status,
            _ => p => null
        };

        /// <summary>
        /// 计划表筛选：条件（值多选/日期范围）+ 关键字
        /// </summary>
        private bool PlanFilter(object obj)
        {
            if (obj is not Plan plan)
            {
                return false;
            }
            foreach (FilterCondition cond in ActiveFilters)
            {
                if (cond.IsDateField)
                {
                    DateTime? d = plan.StartDate;
                    if (d == null)
                    {
                        return false;
                    }
                    if (cond.DateFrom != null && d.Value.Date < cond.DateFrom.Value.Date)
                    {
                        return false;
                    }
                    if (cond.DateTo != null && d.Value.Date > cond.DateTo.Value.Date)
                    {
                        return false;
                    }
                }
                else if (cond.SelectedValues.Count > 0)
                {
                    string value = GetFieldValue(cond.Field)(plan);
                    if (!cond.SelectedValues.Contains(value ?? ""))
                    {
                        return false;
                    }
                }
            }

            if (string.IsNullOrWhiteSpace(Keyword))
            {
                return true;
            }
            string kw = Keyword.Trim();
            return Contains(plan.ModelName, kw) || Contains(plan.JobNo, kw)
                || Contains(plan.TestItem, kw) || Contains(plan.Owner, kw)
                || Contains(plan.Product, kw) || Contains(plan.Customer, kw);
        }

        /// <summary>
        /// 领退表筛选：关键字
        /// </summary>
        private bool RequisitionFilter(object obj)
        {
            if (obj is not Requisition req)
            {
                return false;
            }
            if (string.IsNullOrWhiteSpace(ReqKeyword))
            {
                return true;
            }
            string kw = ReqKeyword.Trim();
            return Contains(req.ModelName, kw) || Contains(req.RequisitionNo, kw)
                || Contains(req.WorkOrder, kw) || Contains(req.ReturnRtOrder, kw)
                || Contains(req.SN, kw);
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
        /// 机种联动：输入机种名称后自动带出产品别/客户别（还原计划表公式关系，仅填充空字段）
        /// </summary>
        public void AutoFillByModel(Plan plan)
        {
            if (string.IsNullOrWhiteSpace(plan.ModelName))
            {
                return;
            }
            ModelMapping mapping = _adminService.FindModelMapping(plan.ModelName);
            if (mapping == null)
            {
                return;
            }
            if (string.IsNullOrWhiteSpace(plan.Product) && mapping.Product != null)
            {
                plan.Product = mapping.Product;
            }
            if (string.IsNullOrWhiteSpace(plan.Customer) && mapping.Customer != null)
            {
                plan.Customer = mapping.Customer;
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
            CommandManager.InvalidateRequerySuggested();
        }

        /* ###############################  增删改  ################################ */

        private void AddRequisition()
        {
            if (!CanEdit)
            {
                StatusMessage = "无新增权限，请登录相应账号";
                return;
            }
            Views.WindowRequisitionEdit editWindow = new(_db, _permission, _adminService, _excelService, null, NeedsReview)
            {
            };
            if (editWindow.ShowDialog() != true)
            {
                return;
            }
            if (NeedsReview)
            {
                _reviewService.SubmitPlanRequest("新增", editWindow.PlanResult, null, _permission.CurrentUser);
                _reviewService.SubmitRequisitionRequest("新增", editWindow.RequisitionResult, null, _permission.CurrentUser);
                StatusMessage = "新增请求已提交审核，待审核员通过后生效";
                _ = System.Windows.MessageBox.Show("新增请求已提交审核，待审核员通过后生效。", "提交成功");
            }
            else
            {
                Refresh();
                StatusMessage = "新增成功";
            }
        }

        private void AddPlan()
        {
            if (!CanEdit)
            {
                StatusMessage = "无新增权限，请登录相应账号";
                return;
            }
            Views.WindowPlanDirectEdit editWindow = new(_db, _permission, _adminService, _excelService, null, NeedsReview)
            {
            };
            if (editWindow.ShowDialog() != true)
            {
                return;
            }
            if (NeedsReview)
            {
                _reviewService.SubmitPlanRequest("新增", editWindow.PlanResult, null, _permission.CurrentUser);
                StatusMessage = "新增请求已提交审核，待审核员通过后生效";
                _ = System.Windows.MessageBox.Show("新增请求已提交审核，待审核员通过后生效。", "提交成功");
            }
            else
            {
                Refresh();
                StatusMessage = "新增成功";
            }
        }

        private void EditRequisition()
        {
            if (!CanEdit || SelectedRequisition == null)
            {
                return;
            }
            Views.WindowRequisitionEdit editWindow = new(_db, _permission, _adminService, _excelService, SelectedRequisition, NeedsReview)
            {
            };
            if (editWindow.ShowDialog() != true)
            {
                return;
            }
            if (NeedsReview)
            {
                _reviewService.SubmitRequisitionRequest("编辑", editWindow.RequisitionResult, SelectedRequisition.Id, _permission.CurrentUser);
                StatusMessage = "编辑请求已提交审核，待审核员通过后生效";
                _ = System.Windows.MessageBox.Show("编辑请求已提交审核，待审核员通过后生效。", "提交成功");
            }
            else
            {
                Refresh();
                StatusMessage = "编辑成功";
            }
        }

        private void EditPlan()
        {
            if (!CanEdit || SelectedPlan == null)
            {
                return;
            }
            Views.WindowPlanDirectEdit editWindow = new(_db, _permission, _adminService, _excelService, SelectedPlan, NeedsReview)
            {
            };
            if (editWindow.ShowDialog() != true)
            {
                return;
            }
            if (NeedsReview)
            {
                _reviewService.SubmitPlanRequest("编辑", editWindow.PlanResult, SelectedPlan.Id, _permission.CurrentUser);
                StatusMessage = "编辑请求已提交审核，待审核员通过后生效";
                _ = System.Windows.MessageBox.Show("编辑请求已提交审核，待审核员通过后生效。", "提交成功");
            }
            else
            {
                Refresh();
                StatusMessage = "编辑成功";
            }
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
                $"确认删除该领退记录？（暂存，提交后生效）\n領料單据號: {reqNo}  机种: {reqModel}",
                "删除确认", System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Warning)
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
                $"确认删除该计划记录？（暂存，提交后生效）\n工作編號: {jobNo}  机种: {planModel}",
                "删除确认", System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Warning)
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
                StatusMessage = $"SN文件不存在: {path}";
                _ = System.Windows.MessageBox.Show($"SN文件不存在:\n{path}", "提示");
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
            if (System.Windows.MessageBox.Show("确认丢弃所有未提交的修改？",
                "丢弃确认", System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Question)
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
                _ = System.Windows.MessageBox.Show($"保存失败:\n{ex.Message}", "错误");
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
                StatusMessage = $"领退表已导出: {file}";
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "导出领退表失败");
                _ = System.Windows.MessageBox.Show($"导出领退表失败:\n{ex.Message}", "错误");
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
                _ = System.Windows.MessageBox.Show($"导出计划表失败:\n{ex.Message}", "错误");
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
            if (System.Windows.MessageBox.Show($"确认清空全部计划数据（计划表{Plans.Count}条/领退表{Requisitions.Count}条）？此操作不可恢复！",
                "清空确认", System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Warning)
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
