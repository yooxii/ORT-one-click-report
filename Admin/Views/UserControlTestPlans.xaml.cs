using Microsoft.Extensions.DependencyInjection;
using NLog;
using ORT一键报告.Main.Views;
using ORT一键报告.Models;
using ORT一键报告.Services;
using ORT一键报告.Utils;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace ORT一键报告.Admin.Views
{
    /// <summary>
    /// "待确认差异"列表的一行：展示某个机种计划里与共享模板不同的字段，以及索引发现的其他写法
    /// </summary>
    public class PlanItemDiffRow
    {
        public TestPlanItem Item { get; set; }
        public PlanItemTemplate Template { get; set; }
        public string PlanTitle { get; set; }
        public string TestItemName { get; set; }
        public string OverrideFields { get; set; }
        public string ModelValue { get; set; }
        public string TemplateValue { get; set; }
        public string Variants { get; set; }
    }

    /// <summary>
    /// 测试计划管理：机种+阶段计划的维护（含拖动排序）、共享测试项模板维护、
    /// 计划索引的执行与进度、以及"待人工确认差异"的采纳/保留。
    /// </summary>
    public partial class UserControlTestPlans : UserControl
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly TestPlanService _planService;
        private readonly PlanIndexService _indexService;
        private readonly AppSettingsService _settings;
        private readonly IPermissionService _permission;

        private readonly ObservableCollection<TestPlan> _plans = [];
        private readonly ObservableCollection<TestPlanItem> _items = [];
        private readonly ObservableCollection<PlanItemTemplate> _templates = [];
        private readonly ObservableCollection<PlanItemDiffRow> _diffs = [];

        private readonly DispatcherTimer _progressTimer = new() { Interval = TimeSpan.FromSeconds(3) };
        private bool _loading;
        private bool _subscribed;

        public UserControlTestPlans()
        {
            InitializeComponent();
            _planService = App.ServiceProvider.GetRequiredService<TestPlanService>();
            _indexService = App.ServiceProvider.GetRequiredService<PlanIndexService>();
            _settings = App.ServiceProvider.GetRequiredService<AppSettingsService>();
            _permission = App.ServiceProvider.GetRequiredService<IPermissionService>();

            dg_plans.ItemsSource = _plans;
            dg_items.ItemsSource = _items;
            dg_templates.ItemsSource = _templates;
            dg_diffs.ItemsSource = _diffs;

            DataGridRowDrag.Enable(dg_items, MoveItemTo);

            // 定时器与 Tick 只接一次（控件会随 Tab 切换反复 Loaded/Unloaded，重复接线会累积处理函数）
            _progressTimer.Tick += (s, args) => RefreshProgress();

            Loaded += OnLoaded;
            Unloaded += OnUnloaded;
        }

        /* ###############################  生命周期  ################################ */

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            if (!_subscribed)
            {
                _indexService.Changed += OnIndexChanged;
                _subscribed = true;
            }
            _progressTimer.Start();
            RefreshPermissions();
            ReloadAll();
            RefreshProgress();
        }

        private void OnUnloaded(object sender, RoutedEventArgs e)
        {
            if (_subscribed)
            {
                _indexService.Changed -= OnIndexChanged;
                _subscribed = false;
            }
            _progressTimer.Stop();
        }

        /// <summary>管理员才能改；技术员可查看（权限沿用计划编辑）</summary>
        private void RefreshPermissions()
        {
            bool canEdit = _permission.Can("admin.manage") || _permission.Can("plan.edit");
            panel_editor.IsEnabled = canEdit;
            btn_index.IsEnabled = canEdit;
            btn_reindex.IsEnabled = canEdit;
            btn_merge.IsEnabled = canEdit;
            btn_syncItems.IsEnabled = canEdit;
            btn_stop.IsEnabled = _indexService.IsRunning;
            txt_root.Text = string.IsNullOrWhiteSpace(_settings.ReportDir)
                ? LanguageService.Get("PlanIndex_NoRoot")
                : $"{LanguageService.Get("PlanIndex_ReportRoot")}: {_settings.ReportDir}";
        }

        /* ###############################  数据加载  ################################ */

        private void ReloadAll()
        {
            _loading = true;
            try
            {
                long selectedPlanId = SelectedPlan?.Id ?? 0;
                _plans.Clear();
                foreach (TestPlan plan in _planService.GetPlans())
                {
                    _plans.Add(plan);
                }
                if (_plans.Count > 0)
                {
                    dg_plans.SelectedItem = _plans.FirstOrDefault(p => p.Id == selectedPlanId) ?? _plans[0];
                }
                ReloadTemplates();
                ReloadDiffs();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "加载测试计划失败");
                ShowMessage(string.Format(LanguageService.Get("Plans_LoadFailed"), ex.Message), true);
            }
            finally
            {
                _loading = false;
            }
            ReloadItems();
        }

        private void ReloadItems()
        {
            _items.Clear();
            TestPlan plan = SelectedPlan;
            if (plan == null)
            {
                ClearItemEditor();
                return;
            }
            try
            {
                foreach (TestPlanItem item in _planService.GetItems(plan.Id))
                {
                    _items.Add(item);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "加载计划明细失败");
            }
            if (_items.Count > 0)
            {
                dg_items.SelectedItem = _items[0];
            }
            else
            {
                ClearItemEditor();
            }
        }

        private void ReloadTemplates()
        {
            long selectedId = SelectedTemplate?.Id ?? 0;
            _templates.Clear();
            foreach (PlanItemTemplate template in _planService.GetTemplates())
            {
                _templates.Add(template);
            }
            if (_templates.Count > 0)
            {
                dg_templates.SelectedItem = _templates.FirstOrDefault(t => t.Id == selectedId) ?? _templates[0];
            }
        }

        private void ReloadDiffs()
        {
            _diffs.Clear();
            try
            {
                Dictionary<long, string> titles = _planService.GetPlanTitles();
                foreach (TestPlanItem item in _planService.GetPendingDiffs())
                {
                    PlanItemTemplate template = item.Template;
                    _diffs.Add(new PlanItemDiffRow
                    {
                        Item = item,
                        Template = template,
                        PlanTitle = titles.TryGetValue(item.PlanId, out string title) ? title : item.PlanId.ToString(),
                        TestItemName = item.TestItemName,
                        OverrideFields = item.OverriddenFieldsDisplay,
                        ModelValue = BuildFieldText(item, false),
                        TemplateValue = BuildFieldText(item, true),
                        Variants = item.SourceVariants
                    });
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "加载待确认差异失败");
            }
        }

        private TestPlan SelectedPlan => dg_plans.SelectedItem as TestPlan;
        private TestPlanItem SelectedItem => dg_items.SelectedItem as TestPlanItem;
        private PlanItemTemplate SelectedTemplate => dg_templates.SelectedItem as PlanItemTemplate;
        private PlanItemDiffRow SelectedDiff => dg_diffs.SelectedItem as PlanItemDiffRow;

        /* ###############################  明细编辑  ################################ */

        private void Dg_Plans_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!_loading)
            {
                ReloadItems();
            }
        }

        private void Dg_Items_SelectionChanged(object sender, SelectionChangedEventArgs e) => FillItemEditor();

        private void FillItemEditor()
        {
            TestPlanItem item = SelectedItem;
            if (item == null)
            {
                ClearItemEditor();
                return;
            }
            txt_itemName.Text = item.TestItemName ?? "";
            txt_itemCategory.Text = item.Category ?? "";
            txt_itemSampling.Text = item.EffectiveSamplingPlan ?? "";
            txt_itemCondition.Text = item.EffectiveTestCondition ?? "";
            txt_itemCriterion.Text = item.EffectivePassCriterion ?? "";
            txt_itemRemark.Text = item.EffectiveRemark ?? "";
            txt_itemPeriod.Text = item.EffectivePeriod ?? "";
            PlanItemTemplate template = item.Template;
            txt_tplName.Text = template?.TestItemName ?? "";
            txt_tplSampling.Text = template?.SamplingPlan ?? "";
            txt_tplCondition.Text = template?.TestCondition ?? "";
            txt_tplCriterion.Text = template?.PassCriterion ?? "";
            txt_tplRemark.Text = template?.Remark ?? "";
        }

        private void ClearItemEditor()
        {
            foreach (TextBox box in new[] { txt_itemName, txt_itemCategory, txt_itemSampling, txt_itemCondition,
                txt_itemCriterion, txt_itemRemark, txt_itemPeriod, txt_tplName, txt_tplSampling, txt_tplCondition,
                txt_tplCriterion, txt_tplRemark })
            {
                box.Text = "";
            }
        }

        /// <summary>
        /// 保存明细：把编辑器里的值写回（与模板相同的字段由服务自动清空，保持"只存差异"）
        /// </summary>
        private void Btn_SaveItem_Click(object sender, RoutedEventArgs e)
        {
            TestPlanItem item = SelectedItem;
            TestPlan plan = SelectedPlan;
            if (item == null || plan == null)
            {
                ShowMessage(LanguageService.Get("PlanIndex_Msg_NoItemSelected"), true);
                return;
            }
            if (string.IsNullOrWhiteSpace(txt_itemName.Text))
            {
                ShowMessage(LanguageService.Get("PlanIndex_Msg_ItemNameRequired"), true);
                txt_itemName.Focus();
                return;
            }
            try
            {
                item.TestItemName = txt_itemName.Text.Trim();
                item.Category = txt_itemCategory.Text.Trim();
                item.TemplateId ??= FindTemplateIdByName(item.TestItemName);
                item.Template ??= _templates.FirstOrDefault(t => t.Id == item.TemplateId);
                item.SamplingPlan = txt_itemSampling.Text;
                item.TestCondition = txt_itemCondition.Text;
                item.PassCriterion = txt_itemCriterion.Text;
                item.Remark = txt_itemRemark.Text;
                item.Period = txt_itemPeriod.Text.Trim();
                _planService.SaveItem(item, CurrentUser);
                int index = _items.IndexOf(item);
                ReloadItems();
                if (index >= 0 && index < _items.Count)
                {
                    dg_items.SelectedItem = _items[index];
                }
                ReloadDiffs();
                ShowMessage(LanguageService.Get("PlanIndex_Msg_Saved"), false);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "保存计划明细失败");
                ShowMessage(ex.Message, true);
            }
        }

        private long? FindTemplateIdByName(string name)
        {
            string key = PlanIndexService.NameKey(name);
            return _templates.FirstOrDefault(t => PlanIndexService.NameKey(t.TestItemName) == key)?.Id;
        }

        private void Btn_AddItem_Click(object sender, RoutedEventArgs e)
        {
            TestPlan plan = SelectedPlan;
            if (plan == null)
            {
                ShowMessage(LanguageService.Get("PlanIndex_Msg_NoPlanSelected"), true);
                return;
            }
            WindowAdminInput dialog = new(LanguageService.Get("PlanIndex_AddItem"),
                (LanguageService.Get("PlanIndex_ItemName"), "", false),
                (LanguageService.Get("PlanIndex_Category"), "RELIABILITY TEST", false));
            if (dialog.ShowDialog() != true)
            {
                return;
            }
            string name = dialog.Values[0];
            if (string.IsNullOrWhiteSpace(name))
            {
                ShowMessage(LanguageService.Get("PlanIndex_Msg_ItemNameRequired"), true);
                return;
            }
            try
            {
                PlanItemTemplate template = _templates.FirstOrDefault(t => PlanIndexService.NameKey(t.TestItemName) == PlanIndexService.NameKey(name));
                TestPlanItem item = new()
                {
                    PlanId = plan.Id,
                    TestItemName = name.Trim(),
                    Category = string.IsNullOrWhiteSpace(dialog.Values[1]) ? template?.Category : dialog.Values[1].Trim(),
                    TemplateId = template?.Id,
                    Template = template
                };
                _planService.SaveItem(item, CurrentUser);
                ReloadAll();
                dg_items.SelectedItem = _items.FirstOrDefault(i => i.Id == item.Id);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "新增计划明细失败");
                ShowMessage(ex.Message, true);
            }
        }

        private void Btn_DeleteItem_Click(object sender, RoutedEventArgs e)
        {
            TestPlanItem item = SelectedItem;
            if (item == null)
            {
                ShowMessage(LanguageService.Get("PlanIndex_Msg_NoItemSelected"), true);
                return;
            }
            if (MessageBox.Show(LanguageService.Get("PlanIndex_Msg_DeleteItem"), LanguageService.Get("Cap_Info"),
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            {
                return;
            }
            _planService.DeleteItem(item.Id);
            ReloadAll();
        }

        private void Btn_MoveUp_Click(object sender, RoutedEventArgs e) => MoveItem(-1);

        private void Btn_MoveDown_Click(object sender, RoutedEventArgs e) => MoveItem(1);

        /// <summary>上移/下移一行（并立即保存顺序）</summary>
        private void MoveItem(int delta)
        {
            TestPlanItem item = SelectedItem;
            int index = item == null ? -1 : _items.IndexOf(item);
            if (index < 0 || index + delta < 0 || index + delta >= _items.Count)
            {
                return;
            }
            _items.Move(index, index + delta);
            PersistOrder();
            dg_items.SelectedItem = item;
        }

        /// <summary>拖动排序：把 source 行移到 target 行的位置</summary>
        private void MoveItemTo(object source, object target)
        {
            if (source is not TestPlanItem from || target is not TestPlanItem to)
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
            PersistOrder();
            dg_items.SelectedItem = from;
        }

        /// <summary>把界面上的顺序写库（顺序号从 1 开始）</summary>
        private void PersistOrder()
        {
            TestPlan plan = SelectedPlan;
            if (plan == null)
            {
                return;
            }
            for (int i = 0; i < _items.Count; i++)
            {
                _items[i].OrderNo = i + 1;
            }
            _planService.SaveOrder(plan.Id, _items.Select(i => i.Id).Where(id => id > 0));
        }

        private void Btn_AdoptTemplate_Click(object sender, RoutedEventArgs e)
        {
            TestPlanItem item = SelectedItem;
            if (item == null)
            {
                ShowMessage(LanguageService.Get("PlanIndex_Msg_NoItemSelected"), true);
                return;
            }
            _planService.AdoptTemplate(item.Id, CurrentUser);
            ReloadAll();
        }

        private void Btn_ConfirmItem_Click(object sender, RoutedEventArgs e) => SetItemConfirmed(true);

        private void Btn_UnconfirmItem_Click(object sender, RoutedEventArgs e) => SetItemConfirmed(false);

        private void SetItemConfirmed(bool confirmed)
        {
            TestPlanItem item = SelectedItem;
            if (item == null)
            {
                ShowMessage(LanguageService.Get("PlanIndex_Msg_NoItemSelected"), true);
                return;
            }
            _planService.SetConfirmed(item.Id, confirmed, CurrentUser);
            ReloadAll();
        }

        /* ###############################  计划  ################################ */

        private void Btn_AddPlan_Click(object sender, RoutedEventArgs e)
        {
            WindowAdminInput dialog = new(LanguageService.Get("PlanIndex_AddPlan"),
                (LanguageService.Get("Plans_ModelName"), "", false),
                (LanguageService.Get("Plans_Stage"), PlanStage.MP, false));
            if (dialog.ShowDialog() != true)
            {
                return;
            }
            if (string.IsNullOrWhiteSpace(dialog.Values[0]))
            {
                ShowMessage(LanguageService.Get("PlanIndex_Msg_PlanNameRequired"), true);
                return;
            }
            try
            {
                TestPlan plan = new()
                {
                    ModelName = dialog.Values[0].Trim(),
                    Stage = PlanStage.Normalize(dialog.Values[1]),
                    Source = "Manual"
                };
                _planService.SavePlan(plan, CurrentUser);
                ReloadAll();
                dg_plans.SelectedItem = _plans.FirstOrDefault(p => p.Id == plan.Id);
            }
            catch (Exception ex)
            {
                ShowMessage(ex.Message, true);
            }
        }

        private void Btn_DeletePlan_Click(object sender, RoutedEventArgs e)
        {
            TestPlan plan = SelectedPlan;
            if (plan == null)
            {
                ShowMessage(LanguageService.Get("PlanIndex_Msg_NoPlanSelected"), true);
                return;
            }
            string title = $"{plan.ModelName} {plan.StageDisplay}";
            if (MessageBox.Show(string.Format(LanguageService.Get("PlanIndex_Msg_DeletePlanFormat"), title),
                LanguageService.Get("Cap_Info"), MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            {
                return;
            }
            _planService.DeletePlan(plan.Id);
            ReloadAll();
        }

        /* ###############################  模板  ################################ */

        private void Dg_Templates_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            PlanItemTemplate template = SelectedTemplate;
            txt_tplEditSampling.Text = template?.SamplingPlan ?? "";
            txt_tplEditCondition.Text = template?.TestCondition ?? "";
            txt_tplEditCriterion.Text = template?.PassCriterion ?? "";
            txt_tplEditRemark.Text = template?.Remark ?? "";
        }

        private void Btn_AddTemplate_Click(object sender, RoutedEventArgs e)
        {
            WindowAdminInput dialog = new(LanguageService.Get("PlanIndex_AddTemplate"),
                (LanguageService.Get("PlanIndex_ItemName"), "", false),
                (LanguageService.Get("PlanIndex_Category"), "RELIABILITY TEST", false),
                (LanguageService.Get("PlanIndex_Period"), "24", false));
            if (dialog.ShowDialog() != true || string.IsNullOrWhiteSpace(dialog.Values[0]))
            {
                return;
            }
            try
            {
                PlanItemTemplate template = new()
                {
                    TestItemName = dialog.Values[0].Trim(),
                    Category = dialog.Values[1]?.Trim(),
                    Period = dialog.Values[2]?.Trim(),
                    IsManual = true
                };
                _planService.SaveTemplate(template, true, CurrentUser);
                _planService.EnsureTestItemRegistered(template.TestItemName, template.Period, CurrentUser);
                ReloadTemplates();
                dg_templates.SelectedItem = _templates.FirstOrDefault(t => t.Id == template.Id);
            }
            catch (Exception ex)
            {
                ShowMessage(ex.Message, true);
            }
        }

        private void Btn_SaveTemplate_Click(object sender, RoutedEventArgs e)
        {
            PlanItemTemplate template = SelectedTemplate;
            if (template == null)
            {
                ShowMessage(LanguageService.Get("PlanIndex_Msg_NoTemplateSelected"), true);
                return;
            }
            try
            {
                template.SamplingPlan = txt_tplEditSampling.Text;
                template.TestCondition = txt_tplEditCondition.Text;
                template.PassCriterion = txt_tplEditCriterion.Text;
                template.Remark = txt_tplEditRemark.Text;
                _planService.SaveTemplate(template, true, CurrentUser);
                ReloadTemplates();
                ShowMessage(LanguageService.Get("PlanIndex_Msg_Saved"), false);
            }
            catch (Exception ex)
            {
                ShowMessage(ex.Message, true);
            }
        }

        private void Btn_DeleteTemplate_Click(object sender, RoutedEventArgs e)
        {
            PlanItemTemplate template = SelectedTemplate;
            if (template == null)
            {
                ShowMessage(LanguageService.Get("PlanIndex_Msg_NoTemplateSelected"), true);
                return;
            }
            if (MessageBox.Show(LanguageService.Get("PlanIndex_Msg_DeleteTemplate"), LanguageService.Get("Cap_Info"),
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            {
                return;
            }
            _planService.DeleteTemplate(template.Id);
            ReloadTemplates();
        }

        private void Btn_SyncTestItems_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                int added = _planService.SyncTestItemsCatalog(CurrentUser);
                ShowMessage(added > 0
                    ? string.Format(LanguageService.Get("PlanIndex_Msg_SyncResultFormat"), added)
                    : LanguageService.Get("PlanIndex_Msg_SyncNone"), false);
            }
            catch (Exception ex)
            {
                ShowMessage(ex.Message, true);
            }
        }

        /* ###############################  待确认差异  ################################ */

        private void Btn_DiffAdopt_Click(object sender, RoutedEventArgs e)
        {
            PlanItemDiffRow row = SelectedDiff;
            if (row == null)
            {
                ShowMessage(LanguageService.Get("PlanIndex_Msg_NoItemSelected"), true);
                return;
            }
            _planService.AdoptTemplate(row.Item.Id, CurrentUser);
            ReloadAll();
        }

        private void Btn_DiffConfirm_Click(object sender, RoutedEventArgs e)
        {
            PlanItemDiffRow row = SelectedDiff;
            if (row == null)
            {
                ShowMessage(LanguageService.Get("PlanIndex_Msg_NoItemSelected"), true);
                return;
            }
            _planService.SetConfirmed(row.Item.Id, true, CurrentUser);
            ReloadAll();
        }

        private void Btn_RefreshDiffs_Click(object sender, RoutedEventArgs e) => ReloadDiffs();

        /* ###############################  索引执行  ################################ */

        private async void Btn_Index_Click(object sender, RoutedEventArgs e)
        {
            string root = _settings.ReportDir;
            if (string.IsNullOrWhiteSpace(root))
            {
                ShowMessage(LanguageService.Get("PlanIndex_NoRoot"), true);
                return;
            }
            await RunIndexAsync(root, false);
        }

        private async void Btn_Reindex_Click(object sender, RoutedEventArgs e)
        {
            string root = _settings.ReportDir;
            if (string.IsNullOrWhiteSpace(root))
            {
                ShowMessage(LanguageService.Get("PlanIndex_NoRoot"), true);
                return;
            }
            if (MessageBox.Show(LanguageService.Get("PlanIndex_Msg_ReindexConfirm"), LanguageService.Get("Cap_Info"),
                MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes)
            {
                return;
            }
            await RunIndexAsync(root, true);
        }

        private async System.Threading.Tasks.Task RunIndexAsync(string root, bool forceRebuild)
        {
            try
            {
                btn_index.IsEnabled = false;
                ShowMessage(LanguageService.Get("PlanIndex_Msg_IndexRunning"), false);
                PlanIndexRunResult result = await _indexService.RunAsync(root, CurrentUser, forceRebuild);
                RefreshProgress();
                ReloadAll();
                if (!result.Started)
                {
                    ShowMessage(string.IsNullOrWhiteSpace(result.Message) ? LanguageService.Get("PlanIndex_Msg_OtherClient") : result.Message, false);
                }
                else if (result.Completed)
                {
                    ShowMessage(string.Format(LanguageService.Get("PlanIndex_Msg_IndexDoneFormat"), result.Message), false);
                }
                else
                {
                    ShowMessage(result.Message, false);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "执行计划索引失败");
                ShowMessage(string.Format(LanguageService.Get("PlanIndex_Msg_IndexFailedFormat"), ex.Message), true);
            }
            finally
            {
                RefreshPermissions();
            }
        }

        private void Btn_Stop_Click(object sender, RoutedEventArgs e)
        {
            _indexService.RequestStop();
            ShowMessage(LanguageService.Get("PlanIndex_Stop"), false);
        }

        private void Btn_Merge_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                PlanIndexJob job = _indexService.GetLatestJob(_settings.ReportDir);
                if (job == null)
                {
                    ShowMessage(LanguageService.Get("PlanIndex_Msg_NoJob"), true);
                    return;
                }
                PlanIndexMergeResult result = _indexService.Merge(job.Id, CurrentUser);
                job.Message = $"归并完成：模板 {result.TemplateCount} 个、计划 {result.PlanCount} 个、明细 {result.ItemCount} 条、差异 {result.DiffCount} 处";
                ReloadAll();
                ShowMessage(job.Message, false);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "重新归并失败");
                ShowMessage(ex.Message, true);
            }
        }

        private void Btn_Refresh_Click(object sender, RoutedEventArgs e)
        {
            ReloadAll();
            RefreshProgress();
        }

        /// <summary>索引进度变化（服务已切回 UI 线程）</summary>
        private void OnIndexChanged() => RefreshProgress();

        /// <summary>上一次看到的任务状态（用于"跑到完成时自动刷新列表"）</summary>
        private string _lastJobStatus;

        /// <summary>
        /// 刷新进度显示：进度条与状态取自数据库里的任务记录，
        /// 因此别的客户端在跑时这里也能看到进展；任务跑到完成后自动刷新计划/模板/差异列表。
        /// </summary>
        private void RefreshProgress()
        {
            try
            {
                PlanIndexJob job = _indexService.GetLatestJob(_settings.ReportDir);
                btn_stop.IsEnabled = _indexService.IsRunning;
                if (job == null)
                {
                    pb_index.Value = 0;
                    txt_status.Text = "";
                    return;
                }
                int total = Math.Max(job.Total, 1);
                pb_index.Maximum = total;
                pb_index.Value = Math.Min(job.Processed, total);
                StringBuilder text = new();
                text.Append(StatusText(job.Status)).Append("  ").Append(job.ProgressText);
                if (!string.IsNullOrWhiteSpace(job.ClaimedBy) && job.ClaimedBy != PlanIndexService.ClientId)
                {
                    text.Append($"  ← {job.ClaimedBy}");
                }
                if (!string.IsNullOrWhiteSpace(job.Message))
                {
                    text.Append("  ").Append(job.Message);
                }
                txt_status.Text = text.ToString();
                // 任务完成/失败时把列表刷新一遍（索引结果这时才写进计划与模板）
                if (_lastJobStatus != job.Status)
                {
                    bool finished = job.Status is PlanIndexJob.StatusDone or PlanIndexJob.StatusFailed;
                    _lastJobStatus = job.Status;
                    if (finished)
                    {
                        ReloadAll();
                        RefreshPermissions();
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Warn($"刷新索引进度失败: {ex.Message}");
            }
        }

        /// <summary>任务状态的显示文本</summary>
        private static string StatusText(string status) => status switch
        {
            PlanIndexJob.StatusPending => LanguageService.Get("PlanIndex_Status_Pending"),
            PlanIndexJob.StatusRunning => LanguageService.Get("PlanIndex_Status_Running"),
            PlanIndexJob.StatusPaused => LanguageService.Get("PlanIndex_Status_Paused"),
            PlanIndexJob.StatusDone => LanguageService.Get("PlanIndex_Status_Done"),
            PlanIndexJob.StatusFailed => LanguageService.Get("PlanIndex_Status_Failed"),
            _ => status
        };

        /* ###############################  辅助  ################################ */

        private string CurrentUser => _permission.CurrentUser ?? "";

        private void ShowMessage(string message, bool isError)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return;
            }
            ToastService.Show(message, isError ? ToastType.Warning : ToastType.Info);
        }

        /// <summary>
        /// 把差异字段拼成"【字段】内容"的文本；useTemplate=true 时取模板里的对应内容
        /// </summary>
        private static string BuildFieldText(TestPlanItem item, bool useTemplate)
        {
            if (item == null || string.IsNullOrWhiteSpace(item.OverriddenFields))
            {
                return "";
            }
            StringBuilder text = new();
            foreach (string field in item.OverriddenFields.Split([','], StringSplitOptions.RemoveEmptyEntries))
            {
                string label = field switch
                {
                    TestPlanItem.FieldSamplingPlan => LanguageService.Get("PlanIndex_SamplingPlan"),
                    TestPlanItem.FieldTestCondition => LanguageService.Get("PlanIndex_TestCondition"),
                    TestPlanItem.FieldPassCriterion => LanguageService.Get("PlanIndex_PassCriterion"),
                    TestPlanItem.FieldRemark => LanguageService.Get("PlanIndex_Remark"),
                    TestPlanItem.FieldPeriod => LanguageService.Get("PlanIndex_Period"),
                    _ => field
                };
                string value = field switch
                {
                    TestPlanItem.FieldSamplingPlan => useTemplate ? item.Template?.SamplingPlan : item.SamplingPlan,
                    TestPlanItem.FieldTestCondition => useTemplate ? item.Template?.TestCondition : item.TestCondition,
                    TestPlanItem.FieldPassCriterion => useTemplate ? item.Template?.PassCriterion : item.PassCriterion,
                    TestPlanItem.FieldRemark => useTemplate ? item.Template?.Remark : item.Remark,
                    TestPlanItem.FieldPeriod => useTemplate ? item.Template?.Period : item.Period,
                    _ => null
                };
                if (text.Length > 0)
                {
                    text.Append('\n');
                }
                text.Append($"【{label}】").Append(value ?? "");
            }
            return text.ToString();
        }
    }
}
