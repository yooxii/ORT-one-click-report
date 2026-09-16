using Microsoft.Extensions.DependencyInjection;
using NLog;
using ORT一键报告.Admin.Views;
using ORT一键报告.Main.Views;
using ORT一键报告.Models;
using ORT一键报告.Services;
using ORT一键报告.Utils;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace ORT一键报告.Reports.Views
{
    /// <summary>
    /// 报告模板工具：以机种的测试计划为输入，直接生成一份报告模板
    /// （Cover 基本信息 / ORT Plan / Waterfall 序列号与测试安排 / TestStatus）。
    /// 入口：领用和计划的右键（空缺报告的计划记录）、该窗口的菜单与工具栏，以及主窗口工具菜单。
    /// </summary>
    public partial class WindowReportTemplate : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly TestPlanService _planService;
        private readonly ReportTemplateService _templateService;
        private readonly AppSettingsService _settings;
        private readonly DatabaseService _db;
        private readonly IPermissionService _permission;

        private readonly ObservableCollection<ReportTemplateItem> _items = [];
        private bool _loading;
        private string _lastFolder;

        public WindowReportTemplate()
        {
            InitializeComponent();
            _planService = App.ServiceProvider.GetRequiredService<TestPlanService>();
            _templateService = App.ServiceProvider.GetRequiredService<ReportTemplateService>();
            _settings = App.ServiceProvider.GetRequiredService<AppSettingsService>();
            _db = App.ServiceProvider.GetRequiredService<DatabaseService>();
            _permission = App.ServiceProvider.GetRequiredService<IPermissionService>();

            dg_items.ItemsSource = _items;
            DataGridRowDrag.Enable(dg_items, MoveItemTo);

            Loaded += (s, e) => Initialize();
        }

        /* ###############################  初始化  ################################ */

        private void Initialize()
        {
            _loading = true;
            try
            {
                cb_model.ItemsSource = _planService.GetPlans()
                    .Select(p => p.ModelName)
                    .Distinct()
                    .OrderBy(n => n)
                    .ToList();
                txt_output.Text = _settings.ReportDir ?? "";
                dp_start.SelectedDate = DateTime.Today;
                txt_stage.Text = PlanStage.MP;
                txt_period.Text = "WK" + DateTime.Today.Year.ToString("0000").Substring(2)
                    + ReportTemplateService.GetIsoWeek(DateTime.Today).ToString("00");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "初始化报告模板工具失败");
            }
            finally
            {
                _loading = false;
            }
        }

        /// <summary>
        /// 从计划表右键带过来的预填：机种、RT工号、序列号、工令、版本、开始日期，并载入该机种的测试计划
        /// </summary>
        public void PrefillFromPlan(Plan plan, Requisition requisition)
        {
            if (plan == null)
            {
                return;
            }
            _loading = true;
            try
            {
                cb_model.Text = plan.ModelName ?? "";
                txt_jobNo.Text = plan.JobNo ?? "";
                txt_stage.Text = plan.Stage ?? PlanStage.MP;
                // 客户别直接取这条计划记录的"客户别"
                FillCustomer(plan.ModelName, plan.Customer);
                if (plan.StartDate.HasValue)
                {
                    dp_start.SelectedDate = plan.StartDate.Value;
                }
                if (requisition != null)
                {
                    txt_sns.Text = requisition.SN ?? "";
                    txt_workOrder.Text = requisition.WorkOrder ?? "";
                    txt_revision.Text = requisition.Rev ?? "";
                    if (requisition.RequisitionDate.HasValue)
                    {
                        dp_start.SelectedDate = requisition.RequisitionDate.Value;
                    }
                    string week = Report.ParseWeekTag(plan.JobNo) ?? Report.ParseWeekTag(requisition.ReturnRtOrder ?? "");
                    if (!string.IsNullOrWhiteSpace(week))
                    {
                        txt_period.Text = week;
                    }
                }
                else
                {
                    string week = Report.ParseWeekTag(plan.JobNo);
                    if (!string.IsNullOrWhiteSpace(week))
                    {
                        txt_period.Text = week;
                    }
                }
            }
            finally
            {
                _loading = false;
            }
            LoadPlanByModel(plan.ModelName, true);
        }

        /* ###############################  计划载入  ################################ */

        private void Cb_Model_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_loading)
            {
                return;
            }
            if (cb_model.SelectedItem is string model)
            {
                LoadPlanByModel(model, false);
            }
        }

        private void Btn_LoadPlan_Click(object sender, RoutedEventArgs e)
        {
            string model = cb_model.Text?.Trim();
            if (string.IsNullOrWhiteSpace(model))
            {
                ToastService.Show(LanguageService.Get("ReportTemplate_Msg_ModelRequired"), ToastType.Warning);
                return;
            }
            LoadPlanByModel(model, true);
        }

        /// <summary>
        /// 载入机种的测试计划到测试项列表（测试计划文本取自"计划模板 + 机种差异"的执行结果）
        /// </summary>
        /// <param name="overwrite">true=覆盖已调整过的测试项</param>
        private void LoadPlanByModel(string modelName, bool overwrite)
        {
            if (string.IsNullOrWhiteSpace(modelName))
            {
                return;
            }
            try
            {
                // 客户别先按计划表带出（该机种没有测试计划时同样要带出来）
                FillCustomer(modelName, FindPlanByModel(modelName)?.Customer);
                if (!overwrite && _items.Count > 0)
                {
                    return; // 用户已手工调整过，不静默覆盖
                }
                List<TestPlan> plans = _planService.GetPlans()
                    .Where(p => string.Equals(p.ModelName, modelName, StringComparison.CurrentCultureIgnoreCase))
                    .ToList();
                TestPlan plan = plans.FirstOrDefault(p => p.Stage == PlanStage.MP) ?? plans.FirstOrDefault();
                if (plan == null)
                {
                    _items.Clear();
                    ToastService.Show(string.Format(LanguageService.Get("ReportTemplate_Msg_PlanMissingFormat"), modelName), ToastType.Warning);
                    return;
                }
                List<TestPlanItem> planItems = _planService.GetItems(plan.Id);
                _items.Clear();
                Dictionary<string, string> categories = LoadTestCategories();
                foreach (TestPlanItem item in planItems)
                {
                    _items.Add(new ReportTemplateItem
                    {
                        TestItemName = item.TestItemName,
                        Category = ResolveCategory(categories, item.TestItemName, item.Category ?? item.Template?.Category),
                        SamplingPlan = item.EffectiveSamplingPlan,
                        TestCondition = item.EffectiveTestCondition,
                        PassCriterion = item.EffectivePassCriterion,
                        Remark = item.EffectiveRemark,
                        PeriodHours = string.IsNullOrWhiteSpace(item.EffectivePeriod) ? "24" : item.EffectivePeriod
                    });
                }
                txt_stage.Text = plan.Stage;
                AutoSchedule(true);
                if (_items.Count == 0)
                {
                    ToastService.Show(LanguageService.Get("ReportTemplate_Msg_PlanEmpty"), ToastType.Warning);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "载入机种测试计划失败");
                ToastService.Show(string.Format(LanguageService.Get("ReportTemplate_Msg_FailedFormat"), ex.Message), ToastType.Warning);
            }
        }

        /// <summary>
        /// 测试种类表：测试项目名（归一化键）→ 测试种类（测试项目表里维护的值）
        /// </summary>
        private Dictionary<string, string> LoadTestCategories()
        {
            Dictionary<string, string> map = [];
            try
            {
                AdminService admin = App.ServiceProvider.GetRequiredService<AdminService>();
                // 没有测试种类的先在管理端自动归类（关键词 + 历史报告）
                admin.EnsureTestItemCategories();
                foreach (TestItemCatalog entry in admin.GetTestItems())
                {
                    string key = PlanIndexService.NameKey(entry.Name);
                    string category = TestCategories.Normalize(entry.Category);
                    if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(category))
                    {
                        map[key] = category;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"读取测试项目种类失败: {ex.Message}");
            }
            return map;
        }

        /// <summary>
        /// 取测试种类：以"测试项目表"的归类为准（用户可手工归类），
        /// 没有归类时退回计划模板里历史报告的分类，仍没有就归"不确定"
        /// </summary>
        private static string ResolveCategory(Dictionary<string, string> categories, string testItemName, string planCategory)
        {
            string key = PlanIndexService.NameKey(testItemName);
            if (categories != null && !string.IsNullOrEmpty(key) && categories.TryGetValue(key, out string category))
            {
                return category;
            }
            return TestCategories.Normalize(planCategory) ?? TestCategories.Uncertain;
        }

        /// <summary>
        /// 带出客户别：以计划表里的"客户别"为准，计划里没有时再退回机种映射
        /// </summary>
        private void FillCustomer(string modelName, string customer)
        {
            if (!string.IsNullOrWhiteSpace(customer))
            {
                txt_customer.Text = customer.Trim();
                return;
            }
            if (!string.IsNullOrWhiteSpace(txt_customer.Text))
            {
                return;
            }
            try
            {
                ModelMapping mapping = _db.FreeSql.Select<ModelMapping>()
                    .Where(m => m.ModelName == modelName)
                    .First();
                if (mapping != null && !string.IsNullOrWhiteSpace(mapping.Customer))
                {
                    txt_customer.Text = mapping.Customer;
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"查询机种映射失败: {ex.Message}");
            }
        }

        /// <summary>取该机种在计划表里的记录（客户别从计划表带出）</summary>
        private Plan FindPlanByModel(string modelName)
        {
            if (string.IsNullOrWhiteSpace(modelName))
            {
                return null;
            }
            try
            {
                return _db.FreeSql.Select<Plan>()
                    .Where(p => p.ModelName == modelName)
                    .OrderByDescending(p => p.Id)
                    .First();
            }
            catch (Exception ex)
            {
                _logger.Warn($"查询计划表失败: {ex.Message}");
                return null;
            }
        }

        /* ###############################  排期与顺序  ################################ */

        private void Btn_AutoFill_Click(object sender, RoutedEventArgs e)
        {
            AutoSchedule(true);
            ToastService.Show(LanguageService.Get("ReportTemplate_Msg_AutoFilled"), ToastType.Info);
        }

        /// <summary>
        /// 按开始日期与每项测试的试验周期排期（开始日期落在工作日），
        /// 并把表格强制按开始日期排序（表格不允许点表头排序）
        /// </summary>
        private void AutoSchedule(bool refresh)
        {
            DateTime start = dp_start.SelectedDate ?? DateTime.Today;
            ReportTemplateService.Schedule(_items.ToList(), start);
            SortItemsByStart();
            if (refresh)
            {
                dg_items.Items.Refresh();
            }
        }

        /// <summary>
        /// 强制按开始日期升序（没有开始日期的排在最后）；
        /// 排期与手工改日期后都会调用，保证表格顺序始终与排期一致
        /// </summary>
        private void SortItemsByStart()
        {
            List<ReportTemplateItem> sorted = _items
                .OrderBy(i => i.Start ?? DateTime.MaxValue)
                .ThenBy(i => i.End ?? DateTime.MaxValue)
                .ToList();
            for (int target = 0; target < sorted.Count; target++)
            {
                int current = _items.IndexOf(sorted[target]);
                if (current >= 0 && current != target)
                {
                    _items.Move(current, target);
                }
            }
        }

        /// <summary>表格里直接改了开始日期后重新排序</summary>
        private void Dg_Items_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                SortItemsByStart();
                dg_items.Items.Refresh();
            }), System.Windows.Threading.DispatcherPriority.Background);
        }

        private void MoveItem(int delta)
        {
            int index = dg_items.SelectedIndex;
            if (index < 0 || index + delta < 0 || index + delta >= _items.Count)
            {
                return;
            }
            int target = index + delta;
            ReportTemplateItem item = _items[index];
            _items.Move(index, target);
            AutoSchedule(false);
            dg_items.Items.Refresh();
            dg_items.SelectedItem = item;
        }

        private void MoveItemTo(object source, object target)
        {
            if (source is not ReportTemplateItem from || target is not ReportTemplateItem to)
            {
                return;
            }
            int fromIndex = _items.IndexOf(from);
            int toIndex = _items.IndexOf(to);
            if (fromIndex < 0 || toIndex < 0 || fromIndex == toIndex)
            {
                return;
            }
            _items.Move(fromIndex, toIndex);
            AutoSchedule(false);
            dg_items.Items.Refresh();
            dg_items.SelectedItem = from;
        }

        private void Btn_MoveUp_Click(object sender, RoutedEventArgs e) => MoveItem(-1);

        private void Btn_MoveDown_Click(object sender, RoutedEventArgs e) => MoveItem(1);

        /* ###############################  测试项维护  ################################ */

        private void Btn_AddItem_Click(object sender, RoutedEventArgs e)
        {
            // 测试项从"测试项目表"（test_items_catalog）里选，不再手工输入名称
            AdminService admin = App.ServiceProvider.GetRequiredService<AdminService>();
            List<TestItemCatalog> catalog = admin.GetTestItems();
            if (catalog.Count == 0)
            {
                ToastService.Show(LanguageService.Get("ReportTemplate_Msg_CatalogEmpty"), ToastType.Warning);
                return;
            }
            List<TestItemCatalog> picked = WindowTemplateItemPicker.Pick(this, catalog, _items.Select(i => i.TestItemName));
            if (picked.Count == 0)
            {
                return;
            }
            List<PlanItemTemplate> templates = _planService.GetTemplates();
            Dictionary<string, string> categories = LoadTestCategories();
            List<string> added = [];
            foreach (TestItemCatalog entry in picked)
            {
                string name = entry.Name?.Trim();
                string key = PlanIndexService.NameKey(name);
                if (string.IsNullOrWhiteSpace(name) || _items.Any(i => PlanIndexService.NameKey(i.TestItemName) == key))
                {
                    added.Add(name ?? "");
                    continue;
                }
                PlanItemTemplate template = templates.FirstOrDefault(t => PlanIndexService.NameKey(t.TestItemName) == key);
                ReportTemplateItem item = new()
                {
                    TestItemName = name,
                    // 分类（测试种类）以测试项目表的归类为准，计划模板里历史报告的分类作为兜底
                    Category = ResolveCategory(categories, name, template?.Category),
                    SamplingPlan = template?.SamplingPlan,
                    TestCondition = template?.TestCondition,
                    PassCriterion = template?.PassCriterion,
                    Remark = template?.Remark ?? entry.Remark,
                    // 试验周期优先用测试项目表里维护的小时数
                    PeriodHours = !string.IsNullOrWhiteSpace(entry.Period)
                        ? entry.Period.Trim()
                        : (string.IsNullOrWhiteSpace(template?.Period) ? "24" : template.Period)
                };
                _items.Add(item);
            }
            AutoSchedule(true);
            if (added.Count > 0)
            {
                ToastService.Show(string.Format(LanguageService.Get("ReportTemplate_DuplicateSkippedFormat"), string.Join("、", added)), ToastType.Warning);
            }
            if (_items.Count > 0)
            {
                dg_items.SelectedItem = _items[_items.Count - 1];
            }
        }

        private void Btn_RemoveItem_Click(object sender, RoutedEventArgs e)
        {
            ReportTemplateItem item = dg_items.SelectedItem as ReportTemplateItem;
            if (item == null)
            {
                return;
            }
            _items.Remove(item);
            AutoSchedule(true);
        }

        private void Dg_Items_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ReportTemplateItem item = dg_items.SelectedItem as ReportTemplateItem;
            txt_planSampling.Text = item?.SamplingPlan ?? "";
            txt_planCondition.Text = item?.TestCondition ?? "";
            txt_planCriterion.Text = item?.PassCriterion ?? "";
        }

        /* ###############################  生成  ################################ */

        private void Btn_Browse_Click(object sender, RoutedEventArgs e)
        {
            IPathService pathService = App.ServiceProvider.GetRequiredService<IPathService>();
            string dir = pathService.OpenPathDialog(LanguageService.Get("Dlg_SelectDir"), initPath: txt_output.Text, isDir: true);
            if (dir != null)
            {
                txt_output.Text = dir;
            }
        }

        private void Btn_Generate_Click(object sender, RoutedEventArgs e)
        {
            string model = cb_model.Text?.Trim();
            if (string.IsNullOrWhiteSpace(model))
            {
                ToastService.Show(LanguageService.Get("ReportTemplate_Msg_ModelRequired"), ToastType.Warning);
                cb_model.Focus();
                return;
            }
            List<string> sns = (txt_sns.Text ?? "")
                .Split(['\n', '\r'], StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim())
                .Where(s => s.Length > 0)
                .ToList();
            if (sns.Count == 0)
            {
                ToastService.Show(LanguageService.Get("ReportTemplate_Msg_SNRequired"), ToastType.Warning);
                txt_sns.Focus();
                return;
            }
            string output = txt_output.Text?.Trim();
            if (string.IsNullOrWhiteSpace(output) || !Directory.Exists(output))
            {
                ToastService.Show(LanguageService.Get("ReportTemplate_Msg_OutputRequired"), ToastType.Warning);
                return;
            }
            if (_items.Count == 0)
            {
                ToastService.Show(LanguageService.Get("ReportTemplate_Msg_PlanEmpty"), ToastType.Warning);
                return;
            }
            try
            {
                AutoSchedule(true);
                string week = ReportTemplateService.ExtractWeek(txt_period.Text);
                string folderName = ReportTemplateService.BuildFolderName(model, week, txt_jobNo.Text?.Trim());
                string targetFolder = Path.Combine(output, folderName);
                if (Directory.Exists(targetFolder)
                    && MessageBox.Show(LanguageService.Get("ReportTemplate_Msg_FolderExists"), LanguageService.Get("Cap_Info"),
                        MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
                {
                    return;
                }
                ReportTemplateRequest request = new()
                {
                    ModelName = model,
                    Customer = txt_customer.Text?.Trim(),
                    Revision = txt_revision.Text?.Trim(),
                    Stage = string.IsNullOrWhiteSpace(txt_stage.Text) ? PlanStage.MP : txt_stage.Text.Trim(),
                    JobNo = txt_jobNo.Text?.Trim(),
                    Period = txt_period.Text?.Trim(),
                    SerialNumbers = sns,
                    WorkOrder = txt_workOrder.Text?.Trim(),
                    StartDate = dp_start.SelectedDate ?? DateTime.Today,
                    Items = _items.ToList(),
                    OutputRoot = output,
                    CreatedBy = _permission.CurrentUser,
                    PlanNote = null
                };
                // 计划备注（ORT Plan 表末的 Note）从计划里带出
                TestPlan plan = _planService.GetPlans().FirstOrDefault(p =>
                    string.Equals(p.ModelName, model, StringComparison.CurrentCultureIgnoreCase));
                request.PlanNote = plan?.Remark;

                ReportTemplateResult result = _templateService.Generate(request);
                if (result.Ok)
                {
                    _lastFolder = result.Folder;
                    ToastService.Show(string.Format(LanguageService.Get("ReportTemplate_Msg_GeneratedFormat"), result.Folder), ToastType.Info);
                    if (MessageBox.Show(string.Format(LanguageService.Get("ReportTemplate_Msg_GeneratedFormat"), result.Folder),
                        LanguageService.Get("Cap_Success"), MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
                    {
                        OpenFolder(result.Folder);
                    }
                }
                else
                {
                    ToastService.Show(string.Format(LanguageService.Get("ReportTemplate_Msg_FailedFormat"), result.Message), ToastType.Warning);
                    MessageBox.Show(string.Format(LanguageService.Get("ReportTemplate_Msg_FailedFormat"), result.Message),
                        LanguageService.Get("Cap_Error"), MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "生成报告模板失败");
                MessageBox.Show(string.Format(LanguageService.Get("ReportTemplate_Msg_FailedFormat"), ex.Message),
                    LanguageService.Get("Cap_Error"), MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void Btn_OpenFolder_Click(object sender, RoutedEventArgs e)
        {
            string folder = _lastFolder;
            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
            {
                folder = Directory.Exists(txt_output.Text) ? txt_output.Text : null;
            }
            if (folder != null)
            {
                OpenFolder(folder);
            }
        }

        private static void OpenFolder(string folder)
        {
            try
            {
                Process.Start("explorer.exe", $"\"{folder}\"");
            }
            catch (Exception ex)
            {
                LogManager.GetCurrentClassLogger().Warn($"打开目录失败: {ex.Message}");
            }
        }

        private void Btn_Close_Click(object sender, RoutedEventArgs e) => Close();
    }
}
