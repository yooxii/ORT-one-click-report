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
        /// 字段名（机种/测试项目/线别/产品别/客户别/负责人/阶段/状况/领用日期/开始日期）
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
    /// 领退和计划主界面 ViewModel：列表展示、导入/导出、新增/编辑/删除（权限分流：直改或提审）、按需筛选、搜索（防抖）
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
        /// 可筛选的字段定义：(字段名, 是否日期字段)
        /// </summary>
        private static readonly (string Name, bool IsDate)[] FilterFields =
        [
            ("机种", false), ("测试项目", false), ("线别", false), ("产品别", false),
            ("客户别", false), ("负责人", false), ("阶段", false), ("状况", false),
            ("领用日期", true), ("开始日期", true)
        ];

        /// <summary>
        /// 全部记录
        /// </summary>
        public ObservableCollection<Plan> Plans { get; } = [];

        /// <summary>
        /// 带筛选/排序的视图
        /// </summary>
        public ICollectionView PlansView { get; }

        /// <summary>
        /// 当前激活的筛选条件（按需添加）
        /// </summary>
        public ObservableCollection<FilterCondition> ActiveFilters { get; } = [];

        private Plan _selectedPlan;
        /// <summary>
        /// 当前选中的记录
        /// </summary>
        public Plan SelectedPlan { get => _selectedPlan; set => SetProperty(ref _selectedPlan, value); }

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
        /// 是否允许表格内直接编辑（技术员及以上；普通用户走对话框提审）
        /// </summary>
        public bool CanGridEdit => CanEdit && !NeedsReview;

        /// <summary>
        /// 表格是否只读（XAML 绑定用：游客/普通用户只读）
        /// </summary>
        public bool IsGridReadOnly => !CanGridEdit;

        /* ###############################  暂存修改（非实时保存，手动提交）  ################################ */

        /// <summary>
        /// 加载时的原始快照（编辑前的状态，用于变更日志与丢弃还原）
        /// </summary>
        private readonly Dictionary<long, Plan> _originals = [];

        /// <summary>
        /// 已修改的记录（Id>0）
        /// </summary>
        private readonly HashSet<Plan> _pendingModified = [];

        /// <summary>
        /// 新增的行（Id==0）
        /// </summary>
        private readonly List<Plan> _pendingAdded = [];

        /// <summary>
        /// 待删除的记录
        /// </summary>
        private readonly Dictionary<long, Plan> _pendingDeleted = [];

        /// <summary>
        /// 是否有未提交的修改（新增/删除立即感知；已有行的修改以快照对比感知）
        /// </summary>
        public bool HasPendingChanges => _pendingAdded.Count > 0 || _pendingDeleted.Count > 0 || DetectModifiedCount() > 0;

        /// <summary>
        /// 未提交修改的描述（状态栏显示）
        /// </summary>
        public string PendingText => HasPendingChanges
            ? $"未提交修改: 新增{_pendingAdded.Count} 修改{DetectModifiedCount()} 删除{_pendingDeleted.Count}"
            : "无未提交修改";

        /// <summary>
        /// 以快照对比检测已修改的存量行数
        /// </summary>
        private int DetectModifiedCount()
        {
            int count = 0;
            foreach (Plan plan in Plans.Where(p => p.Id > 0))
            {
                if (_originals.TryGetValue(plan.Id, out Plan before)
                    && JsonConvert.SerializeObject(plan) != JsonConvert.SerializeObject(before))
                {
                    count++;
                }
            }
            return count;
        }

        /* ###############################  字典（编辑选择框与校验用）  ################################ */

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
        /// 搜索关键字（输入防抖 250ms，避免逐字符触发全量刷新卡顿）
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

        public PlansViewModel(DatabaseService db, PlanExcelService excelService, IPathService pathService, IPermissionService permission, ReviewService reviewService, AdminService adminService)
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

        private RelayCommand _importRequisitionCommand;
        public ICommand ImportRequisitionCommand => _importRequisitionCommand ??= new RelayCommand(ImportRequisition);

        private RelayCommand _importScheduleCommand;
        public ICommand ImportScheduleCommand => _importScheduleCommand ??= new RelayCommand(ImportSchedule);

        private RelayCommand _exportRequisitionCommand;
        public ICommand ExportRequisitionCommand => _exportRequisitionCommand ??= new RelayCommand(ExportRequisition);

        private RelayCommand _exportScheduleCommand;
        public ICommand ExportScheduleCommand => _exportScheduleCommand ??= new RelayCommand(ExportSchedule);

        private RelayCommand _addCommand;
        public ICommand AddCommand => _addCommand ??= new RelayCommand(AddPlan, () => CanEdit);

        private RelayCommand _editCommand;
        public ICommand EditCommand => _editCommand ??= new RelayCommand(EditPlan, () => CanEdit && SelectedPlan != null);

        private RelayCommand _deleteCommand;
        public ICommand DeleteCommand => _deleteCommand ??= new RelayCommand(DeletePlan, () => CanEdit && SelectedPlan != null);

        private RelayCommand _clearAllCommand;
        public ICommand ClearAllCommand => _clearAllCommand ??= new RelayCommand(ClearAll);

        private RelayCommand _saveChangesCommand;
        /// <summary>
        /// 提交保存暂存的修改（写入数据库并记录变更日志）
        /// </summary>
        public ICommand SaveChangesCommand => _saveChangesCommand ??= new RelayCommand(SaveChanges, () => CanGridEdit && HasPendingChanges);

        private RelayCommand _discardChangesCommand;
        /// <summary>
        /// 丢弃暂存的修改，还原为数据库状态
        /// </summary>
        public ICommand DiscardChangesCommand => _discardChangesCommand ??= new RelayCommand(DiscardChanges, () => CanGridEdit && HasPendingChanges);

        private RelayCommand _addRowCommand;
        /// <summary>
        /// 表格内新增空行（暂存，提交后入库）
        /// </summary>
        public ICommand AddRowCommand => _addRowCommand ??= new RelayCommand(AddRow, () => CanGridEdit);

        private CommunityToolkit.Mvvm.Input.RelayCommand<object> _deleteRowCommand;
        /// <summary>
        /// 标记删除行（暂存，提交后生效；参数为 Plan）
        /// </summary>
        public ICommand DeleteRowCommand => _deleteRowCommand ??= new CommunityToolkit.Mvvm.Input.RelayCommand<object>(DeleteRow, p => CanGridEdit && p is Plan);

        private CommunityToolkit.Mvvm.Input.RelayCommand<object> _removeFilterCommand;
        /// <summary>
        /// 删除指定筛选条件（参数为 FilterCondition）
        /// </summary>
        public ICommand RemoveFilterCommand => _removeFilterCommand ??= new CommunityToolkit.Mvvm.Input.RelayCommand<object>(RemoveFilter);

        private CommunityToolkit.Mvvm.Input.RelayCommand<object> _openSnFileCommand;
        /// <summary>
        /// 打开指定记录的SN文件（列表中以超链接呈现）
        /// </summary>
        public CommunityToolkit.Mvvm.Input.RelayCommand<object> OpenSnFileCommand
            => _openSnFileCommand ??= new CommunityToolkit.Mvvm.Input.RelayCommand<object>(OpenSnFile);

        /* ###############################  功能函数  ################################ */

        /// <summary>
        /// 从数据库重新加载全部记录，并刷新各筛选条件的候选项（丢弃未提交修改）
        /// </summary>
        public void Refresh()
        {
            try
            {
                List<Plan> plans = _db.FreeSql.Select<Plan>().OrderByDescending(p => p.Id).ToList();
                Plans.Clear();
                _originals.Clear();
                _pendingModified.Clear();
                _pendingAdded.Clear();
                _pendingDeleted.Clear();
                foreach (Plan plan in plans)
                {
                    Plans.Add(plan);
                    _originals[plan.Id] = ClonePlan(plan);
                }
                // 刷新各值字段条件的候选项（保留当前勾选）
                foreach (FilterCondition cond in ActiveFilters.Where(c => !c.IsDateField))
                {
                    RefreshConditionOptions(cond);
                }
                LoadCatalogs();
                PlansView.Refresh();
                OnPropertyChanged(nameof(HasPendingChanges));
                OnPropertyChanged(nameof(PendingText));
                StatusMessage = $"共 {Plans.Count} 条记录";
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
            CatalogProducts = _adminService.GetProducts().Select(p => p.Name).ToList();
            CatalogCustomers = _adminService.GetCustomers().Select(c => c.Name).ToList();
            CatalogStages = _adminService.GetStages().Select(s => s.Name).ToList();
            OnPropertyChanged(nameof(CatalogTestItems));
            OnPropertyChanged(nameof(CatalogProducts));
            OnPropertyChanged(nameof(CatalogCustomers));
            OnPropertyChanged(nameof(CatalogStages));
        }

        private static Plan ClonePlan(Plan source)
            => JsonConvert.DeserializeObject<Plan>(JsonConvert.SerializeObject(source));

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

        /// <summary>
        /// 按当前数据重建条件的候选项（去掉已不存在的勾选值）
        /// </summary>
        private void RefreshConditionOptions(FilterCondition cond)
        {
            List<string> options = DistinctOptions(Plans.Select(GetFieldValue(cond.Field)));
            cond.Options = options;
            // 清理已失效的勾选
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
        /// 字段名 -> 取值函数
        /// </summary>
        private static Func<Plan, string> GetFieldValue(string field) => field switch
        {
            "机种" => p => p.ModelName,
            "测试项目" => p => p.TestItem,
            "线别" => p => p.LineNo,
            "产品别" => p => p.Product,
            "客户别" => p => p.Customer,
            "负责人" => p => p.Owner,
            "阶段" => p => p.Stage,
            "状况" => p => p.Status,
            _ => p => null
        };

        private void ImportRequisition()
        {
            if (!_permission.Can("plan.import"))
            {
                StatusMessage = "无导入权限";
                return;
            }
            string file = _pathService.OpenPathDialog("选择领用表(成品領用記錄)");
            if (file == null)
            {
                return;
            }
            try
            {
                (int added, int updated) = _excelService.ImportRequisition(file);
                Refresh();
                StatusMessage = $"领用表导入完成: 新增{added}条, 更新{updated}条";
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "导入领用表失败");
                StatusMessage = $"导入领用表失败: {ex.Message}";
                System.Windows.MessageBox.Show($"导入领用表失败:\n{ex.Message}", "错误");
            }
        }

        private void ImportSchedule()
        {
            if (!_permission.Can("plan.import"))
            {
                StatusMessage = "无导入权限";
                return;
            }
            string file = _pathService.OpenPathDialog("选择计划表(ORT Test Schedule)");
            if (file == null)
            {
                return;
            }
            try
            {
                (int added, int updated, List<string> unmatched) = _excelService.ImportSchedule(file);
                Refresh();
                StatusMessage = $"计划表导入完成: 新增{added}条, 更新{updated}条" + (unmatched.Count > 0 ? $", {unmatched.Count}条未匹配到领用数据" : "");
                if (unmatched.Count > 0)
                {
                    string list = unmatched.Count > 30
                        ? string.Join("\n", unmatched.Take(30)) + $"\n...等共{unmatched.Count}条"
                        : string.Join("\n", unmatched);
                    _ = System.Windows.MessageBox.Show(
                        $"以下 {unmatched.Count} 条计划记录的备注中未找到工令，且工作編號非 Q 开头，未能关联到领用数据：\n\n{list}",
                        "导入提示", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "导入计划表失败");
                StatusMessage = $"导入计划表失败: {ex.Message}";
                System.Windows.MessageBox.Show($"导入计划表失败:\n{ex.Message}", "错误");
            }
        }

        private void ExportRequisition()
        {
            if (!_permission.Can("plan.export"))
            {
                StatusMessage = "无导出权限";
                return;
            }
            string file = _pathService.SavePathDialog("导出领用表", $"{DateTime.Now:yyyy}.成品領用記錄.xlsx");
            if (file == null)
            {
                return;
            }
            try
            {
                _excelService.ExportRequisition(file);
                StatusMessage = $"领用表已导出: {file}";
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "导出领用表失败");
                StatusMessage = $"导出领用表失败: {ex.Message}";
                System.Windows.MessageBox.Show($"导出领用表失败:\n{ex.Message}", "错误");
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
                StatusMessage = $"导出计划表失败: {ex.Message}";
                System.Windows.MessageBox.Show($"导出计划表失败:\n{ex.Message}", "错误");
            }
        }

        private void AddPlan()
        {
            if (!CanEdit)
            {
                StatusMessage = "无新增权限，请登录相应账号";
                return;
            }
            Views.WindowPlanEdit editWindow = new(_db, _permission, null, NeedsReview)
            {
                Owner = System.Windows.Application.Current.MainWindow
            };
            if (editWindow.ShowDialog() != true)
            {
                return;
            }
            if (NeedsReview)
            {
                // 普通用户：不直接写库，提交审核请求
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

        private void EditPlan()
        {
            if (!CanEdit || SelectedPlan == null)
            {
                return;
            }
            Views.WindowPlanEdit editWindow = new(_db, _permission, SelectedPlan, NeedsReview)
            {
                Owner = System.Windows.Application.Current.MainWindow
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

        private void DeletePlan()
        {
            if (SelectedPlan == null)
            {
                return;
            }
            if (!CanEdit)
            {
                StatusMessage = "无删除权限，请登录相应账号";
                return;
            }
            string desc = $"领料单据号: {SelectedPlan.RequisitionNo ?? "-"}  工作編號: {SelectedPlan.JobNo ?? "-"}";
            if (System.Windows.MessageBox.Show($"确认删除该记录？\n{desc}\n機種: {SelectedPlan.ModelName}",
                "删除确认", System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Warning)
                != System.Windows.MessageBoxResult.Yes)
            {
                return;
            }
            if (NeedsReview)
            {
                // 普通用户：提交删除审核请求
                _reviewService.SubmitPlanRequest("删除", SelectedPlan, SelectedPlan.Id, _permission.CurrentUser);
                StatusMessage = "删除请求已提交审核，待审核员通过后生效";
                _ = System.Windows.MessageBox.Show("删除请求已提交审核，待审核员通过后生效。", "提交成功");
                return;
            }
            try
            {
                _db.FreeSql.Delete<Plan>().Where(p => p.Id == SelectedPlan.Id).ExecuteAffrows();
                _logger.Info($"删除计划记录: Id={SelectedPlan.Id}, {desc}");
                Refresh();
                StatusMessage = "删除成功";
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "删除记录失败");
                StatusMessage = $"删除失败: {ex.Message}";
            }
        }

        /// <summary>
        /// 打开SN附件文件
        /// </summary>
        private void OpenSnFile(object parameter)
        {
            if (parameter is not Plan plan || string.IsNullOrWhiteSpace(plan.SnFilePath))
            {
                return;
            }
            string path = _db.ResolveAttachmentPath(plan.SnFilePath);
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

        /// <summary>
        /// 清空全部数据（需二次确认）
        /// </summary>
        private void ClearAll()
        {
            if (!_permission.Can("plan.delete"))
            {
                StatusMessage = "无删除权限";
                return;
            }
            if (Plans.Count == 0)
            {
                StatusMessage = "没有可清空的数据";
                return;
            }
            if (System.Windows.MessageBox.Show($"确认清空全部 {Plans.Count} 条计划数据？此操作不可恢复！",
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

        /* ###############################  表格内编辑：暂存与提交  ################################ */

        /// <summary>
        /// 标记记录已修改（由单元格编辑结束时调用）；Id==0 的行视为新增
        /// </summary>
        public void MarkModified(Plan plan)
        {
            if (plan.Id == 0)
            {
                if (!_pendingAdded.Contains(plan))
                {
                    _pendingAdded.Add(plan);
                }
            }
            else
            {
                _pendingModified.Add(plan);
            }
            OnPropertyChanged(nameof(HasPendingChanges));
            OnPropertyChanged(nameof(PendingText));
            CommandManager.InvalidateRequerySuggested();
        }

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
            bool changed = false;
            if (string.IsNullOrWhiteSpace(plan.Product) && mapping.Product != null)
            {
                plan.Product = mapping.Product;
                changed = true;
            }
            if (string.IsNullOrWhiteSpace(plan.Customer) && mapping.Customer != null)
            {
                plan.Customer = mapping.Customer;
                changed = true;
            }
            if (changed)
            {
                MarkModified(plan);
                _logger.Info($"机种[{plan.ModelName}]自动带出: 产品别={plan.Product}, 客户别={plan.Customer}");
            }
        }

        /// <summary>
        /// 表格内新增空行（暂存）
        /// </summary>
        private void AddRow()
        {
            if (!CanGridEdit)
            {
                return;
            }
            Plan plan = new()
            {
                CreatedBy = _permission.CurrentUser,
                CreatedAt = DateTime.Now,
                UpdatedBy = _permission.CurrentUser,
                UpdatedAt = DateTime.Now
            };
            _pendingAdded.Add(plan);
            Plans.Insert(0, plan);
            OnPropertyChanged(nameof(HasPendingChanges));
            OnPropertyChanged(nameof(PendingText));
            StatusMessage = "已新增空行，填写后请点击提交保存";
        }

        /// <summary>
        /// 标记删除行（暂存，提交后生效）
        /// </summary>
        private void DeleteRow(object parameter)
        {
            if (parameter is not Plan plan || !CanGridEdit)
            {
                return;
            }
            if (plan.Id == 0)
            {
                // 未入库的新增行直接移除
                _pendingAdded.Remove(plan);
                Plans.Remove(plan);
            }
            else
            {
                if (System.Windows.MessageBox.Show(
                    $"确认删除该记录？（暂存，提交后生效）\n工作編號: {plan.JobNo ?? "-"}  机种: {plan.ModelName ?? "-"}",
                    "删除确认", System.Windows.MessageBoxButton.YesNo, System.Windows.MessageBoxImage.Warning)
                    != System.Windows.MessageBoxResult.Yes)
                {
                    return;
                }
                _pendingDeleted[plan.Id] = plan;
                _pendingModified.Remove(plan);
                Plans.Remove(plan);
            }
            OnPropertyChanged(nameof(HasPendingChanges));
            OnPropertyChanged(nameof(PendingText));
            StatusMessage = PendingText;
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
        /// 提交保存：将暂存的新增/修改/删除写入数据库，并为每条变更写入变更日志（便于回滚）
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

                foreach (Plan plan in _pendingAdded)
                {
                    if (string.IsNullOrWhiteSpace(plan.JobNo) && string.IsNullOrWhiteSpace(plan.RequisitionNo))
                    {
                        StatusMessage = "存在未填写工作编号/领料单据号的空行，请补充或删除后再提交";
                        return;
                    }
                    plan.Id = _db.FreeSql.Insert(plan).ExecuteIdentity();
                    WriteChangeLog("新增", plan.Id, $"新增计划 {plan.JobNo ?? plan.RequisitionNo} ({plan.ModelName})", null, plan, op);
                    added++;
                }
                foreach (Plan plan in Plans.Where(p => p.Id > 0))
                {
                    // 以快照对比自动检测修改（兼容单元格编辑与字典选择列等所有编辑方式）
                    if (!_originals.TryGetValue(plan.Id, out Plan before))
                    {
                        continue;
                    }
                    if (JsonConvert.SerializeObject(plan) == JsonConvert.SerializeObject(before))
                    {
                        continue;
                    }
                    plan.UpdatedBy = op;
                    plan.UpdatedAt = DateTime.Now;
                    _db.FreeSql.Update<Plan>().SetSource(plan).Where(p => p.Id == plan.Id).ExecuteAffrows();
                    WriteChangeLog("编辑", plan.Id, $"编辑计划 {plan.JobNo ?? plan.RequisitionNo} ({plan.ModelName})", before, plan, op);
                    modified++;
                }
                foreach (KeyValuePair<long, Plan> kv in _pendingDeleted)
                {
                    _db.FreeSql.Delete<Plan>().Where(p => p.Id == kv.Key).ExecuteAffrows();
                    WriteChangeLog("删除", kv.Key, $"删除计划 {kv.Value.JobNo ?? kv.Value.RequisitionNo} ({kv.Value.ModelName})", kv.Value, null, op);
                    deleted++;
                }

                _logger.Info($"提交保存计划数据: 新增{added} 修改{modified} 删除{deleted} 操作人={op}");
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

        /// <summary>
        /// 写入变更日志（前后快照，便于追溯与回滚）
        /// </summary>
        private void WriteChangeLog(string action, long planId, string summary, Plan before, Plan after, string op)
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

        /// <summary>
        /// 视图筛选：按需添加的条件（值多选/日期范围）+ 关键字
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
                    DateTime? d = cond.Field == "领用日期" ? plan.RequisitionDateValue : plan.StartDateValue;
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
            return Contains(plan.ModelName, kw) || Contains(plan.RequisitionNo, kw)
                || Contains(plan.JobNo, kw) || Contains(plan.TestItem, kw)
                || Contains(plan.Owner, kw) || Contains(plan.SN, kw)
                || Contains(plan.WorkOrder, kw) || Contains(plan.ReturnRtOrder, kw);
        }

        private static bool Contains(string source, string keyword)
            => source?.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0;
    }
}
