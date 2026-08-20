using Microsoft.Extensions.DependencyInjection;
using NLog;
using ORT一键报告.Models;
using ORT一键报告.Services;
using ORT一键报告.Utils;
using System;
using System.Linq;
using System.Windows;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// WindowPlanDirectEdit.xaml 的交互逻辑：计划表直接新增/编辑（第二种情况，非ORT正常领用试验，QRT前缀）。
    /// 必填：测试项目/开始时间/阶段/机种名/备注；自动补全：工作编号 QRT{年月}{编号}、产品别/客户别/负责人/试验时间/结束日期；状态默认 Ongoing。
    /// </summary>
    public partial class WindowPlanDirectEdit : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;
        private readonly IPermissionService _permission;
        private readonly AdminService _admin;
        private readonly PlanExcelService _excelService;
        private readonly Plan _editTarget;
        private readonly bool _submitForReview;

        /// <summary>
        /// 保存/构造的计划记录结果
        /// </summary>
        public Plan PlanResult { get; private set; }

        public WindowPlanDirectEdit(DatabaseService db, IPermissionService permission, AdminService admin,
            PlanExcelService excelService, Plan editTarget = null, bool submitForReview = false)
        {
            InitializeComponent();
            _db = db;
            _permission = permission;
            _admin = admin;
            _excelService = excelService;
            _editTarget = editTarget;
            _submitForReview = submitForReview;

            Title = editTarget == null ? "计划表新增（非领用）" : "计划表编辑";

            cb_testItem.ItemsSource = _admin.GetTestItems().Select(t => t.Name).ToList();
            cb_stage.ItemsSource = _admin.GetStages().Select(s => s.Name).ToList();

            if (editTarget != null)
            {
                LoadFromPlan(editTarget);
            }
        }

        /* ###############################  加载  ################################ */

        private void LoadFromPlan(Plan plan)
        {
            SetCombo(cb_testItem, plan.TestItem);
            SetCombo(cb_stage, plan.Stage);
            dp_startDate.SelectedDate = plan.StartDate;
            txt_model.Text = plan.ModelName;
            txt_jobNo.Text = plan.JobNo;
            txt_sampleSize.Text = plan.SampleSize;
            txt_product.Text = plan.Product;
            txt_customer.Text = plan.Customer;
            txt_owner.Text = plan.Owner;
            txt_testPeriod.Text = plan.TestPeriod;
            dp_endDate.SelectedDate = plan.EndDate;
            txt_remark.Text = plan.Remark;
        }

        private static void SetCombo(System.Windows.Controls.ComboBox combo, string value)
        {
            if (value == null)
            {
                combo.SelectedItem = null;
                return;
            }
            if (combo.ItemsSource is System.Collections.Generic.List<string> list && !list.Contains(value))
            {
                list.Add(value);
            }
            combo.SelectedItem = value;
        }

        /* ###############################  自动补全  ################################ */

        /// <summary>
        /// 根据测试项目自动补全负责人/试验时间/结束日期
        /// </summary>
        private void UpdateAutoPlan()
        {
            string testItem = cb_testItem.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(testItem))
            {
                return;
            }
            TestItemCatalog item = _admin.GetTestItems().FirstOrDefault(t => t.Name == testItem);
            if (item != null)
            {
                txt_owner.Text = item.Owner;
                txt_testPeriod.Text = item.Period;
                if (int.TryParse(item.Period, out int hours) && dp_startDate.SelectedDate is DateTime start)
                {
                    dp_endDate.SelectedDate = start.AddHours(hours);
                }
            }
        }

        /// <summary>
        /// 根据机种名自动补全产品别/客户别
        /// </summary>
        private void UpdateModelMapping()
        {
            string model = txt_model?.Text?.Trim();
            if (string.IsNullOrWhiteSpace(model))
            {
                return;
            }
            ModelMapping mapping = _admin.FindModelMapping(model);
            txt_product.Text = mapping?.Product ?? "";
            txt_customer.Text = mapping?.Customer ?? "";
        }

        /* ###############################  事件函数  ################################ */

        private void Cb_TestItem_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (txt_owner == null)
            {
                return;
            }
            UpdateAutoPlan();
        }

        private void Dp_StartDate_Changed(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (txt_jobNo == null)
            {
                return;
            }
            // 新增时自动生成工作编号 QRT{年月}{编号}
            if (_editTarget == null && dp_startDate.SelectedDate is DateTime start)
            {
                txt_jobNo.Text = _excelService.GenerateJobNo(start, "QRT");
            }
            UpdateAutoPlan();
        }

        private void Btn_Save_Click(object sender, RoutedEventArgs e)
        {
            // 必填校验
            if (cb_testItem.SelectedItem == null)
            {
                _ = MessageBox.Show("請選擇測試項目", "提示");
                return;
            }
            if (dp_startDate.SelectedDate == null)
            {
                _ = MessageBox.Show("請填寫開始時間", "提示");
                return;
            }
            if (cb_stage.SelectedItem == null)
            {
                _ = MessageBox.Show("請選擇階段", "提示");
                return;
            }
            if (string.IsNullOrWhiteSpace(txt_model.Text))
            {
                _ = MessageBox.Show("請填寫機種名稱", "提示");
                return;
            }
            if (string.IsNullOrWhiteSpace(txt_remark.Text))
            {
                _ = MessageBox.Show("請填寫備註", "提示");
                return;
            }

            string jobNo = _editTarget == null
                ? _excelService.GenerateJobNo(dp_startDate.SelectedDate.Value, "QRT")
                : txt_jobNo.Text.Trim();

            // 工作编号格式与唯一性校验
            string jobNoError = PlanValidation.ValidateJobNo(jobNo);
            if (jobNoError != null)
            {
                _ = MessageBox.Show(jobNoError, "格式校验失败");
                return;
            }
            long selfId = _editTarget?.Id ?? 0;
            if (_db.FreeSql.Select<Plan>().Where(p => p.JobNo == jobNo && p.Id != selfId).Any())
            {
                _ = MessageBox.Show($"工作編號 [{jobNo}] 已存在", "提示");
                return;
            }

            Plan plan = _editTarget == null
                ? new Plan { CreatedBy = _permission.CurrentUser, CreatedAt = DateTime.Now }
                : ClonePlan(_editTarget);
            plan.JobNo = jobNo;
            plan.TestItem = cb_testItem.SelectedItem as string;
            plan.StartDate = dp_startDate.SelectedDate;
            plan.Stage = cb_stage.SelectedItem as string;
            plan.ModelName = txt_model.Text.Trim();
            plan.SampleSize = string.IsNullOrWhiteSpace(txt_sampleSize.Text) ? null : txt_sampleSize.Text.Trim();
            plan.Product = string.IsNullOrWhiteSpace(txt_product.Text) ? null : txt_product.Text.Trim();
            plan.Customer = string.IsNullOrWhiteSpace(txt_customer.Text) ? null : txt_customer.Text.Trim();
            plan.Owner = string.IsNullOrWhiteSpace(txt_owner.Text) ? null : txt_owner.Text.Trim();
            plan.TestPeriod = string.IsNullOrWhiteSpace(txt_testPeriod.Text) ? null : txt_testPeriod.Text.Trim();
            plan.EndDate = dp_endDate.SelectedDate;
            plan.Status = plan.Status ?? "Ongoing";
            plan.Remark = txt_remark.Text.Trim();
            plan.UpdatedBy = _permission.CurrentUser;
            plan.UpdatedAt = DateTime.Now;

            PlanResult = plan;

            if (_submitForReview)
            {
                DialogResult = true;
                return;
            }

            try
            {
                if (_editTarget == null)
                {
                    _db.FreeSql.Insert(plan).ExecuteAffrows();
                    _logger.Info($"新增非领用计划: {plan.JobNo}, 机种={plan.ModelName}");
                }
                else
                {
                    _db.FreeSql.Update<Plan>().SetSource(plan).Where(p => p.Id == plan.Id).ExecuteAffrows();
                    _logger.Info($"编辑计划: Id={plan.Id}, {plan.JobNo}");
                }
                DialogResult = true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "保存计划记录失败");
                _ = MessageBox.Show($"保存失败:\n{ex.Message}", "错误");
            }
        }

        private void Btn_Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private static Plan ClonePlan(Plan source)
            => Newtonsoft.Json.JsonConvert.DeserializeObject<Plan>(
                Newtonsoft.Json.JsonConvert.SerializeObject(source));
    }
}
