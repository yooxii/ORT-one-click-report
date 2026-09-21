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
    /// 本对话框只构造结果，不写数据库；由调用方决定暂存或提审。
    /// </summary>
    public partial class WindowPlanDirectEdit : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;
        private readonly IPermissionService _permission;
        private readonly AdminService _admin;
        private readonly PlanExcelService _excelService;
        private readonly Plan _editTarget;

        /// <summary>
        /// 构造的计划记录结果（由调用方处理：暂存或提审）
        /// </summary>
        public Plan PlanResult { get; private set; }

        /// <summary>
        /// 保存成功事件（非模态窗口用）：参数为 (计划结果, 编辑目标Id)。
        /// 调用方订阅后处理暂存/提审逻辑，窗口自身只负责构造结果并关闭。
        /// </summary>
        public event Action<Plan, long> Saved;

        /// <summary>
        /// 「转为领用」事件（仅新增时可用）：参数为当前界面内容组成的计划草稿。
        /// 由调用方打开「领退表新增」并把这些值带过去，最终由领用流程建立 RT 计划。
        /// </summary>
        public event Action<Plan> ConvertToRequisitionRequested;

        public WindowPlanDirectEdit(DatabaseService db, IPermissionService permission, AdminService admin,
            PlanExcelService excelService, Plan editTarget = null)
        {
            InitializeComponent();
            _db = db;
            _permission = permission;
            _admin = admin;
            _excelService = excelService;
            _editTarget = editTarget;

            Title = editTarget == null ? "计划表新增（非领用）" : "计划表编辑";
            // 「转为领用」只在新增时有意义（编辑时已有记录，改走领用要走领退表编辑）
            btn_convertToRequisition.Visibility = editTarget == null ? Visibility.Visible : Visibility.Collapsed;

            cb_testItem.ItemsSource = _admin.GetTestItems().Select(t => t.Name).ToList();
            cb_stage.ItemsSource = _admin.GetStages().Select(s => s.Name).ToList();
            cb_reportStatus.ItemsSource = Models.ReportStatusKind.All.ToList();

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
            SetCombo(cb_reportStatus, plan.ReportStatus);
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
        /// 根据测试项目自动补全负责人/试验时间/结束日期（仅填充空字段，允许手动修改）
        /// </summary>
        private void UpdateAutoPlan()
        {
            string testItem = cb_testItem.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(testItem))
            {
                return;
            }
            TestItemCatalog item = _admin.GetTestItems().FirstOrDefault(t => t.Name == testItem);
            if (item == null)
            {
                return;
            }
            if (string.IsNullOrWhiteSpace(txt_owner.Text))
            {
                txt_owner.Text = item.Owner;
            }
            if (string.IsNullOrWhiteSpace(txt_testPeriod.Text))
            {
                txt_testPeriod.Text = item.Period;
            }
            if (dp_endDate.SelectedDate == null && int.TryParse(item.Period, out int hours) && dp_startDate.SelectedDate is DateTime start)
            {
                dp_endDate.SelectedDate = start.AddHours(hours);
            }
        }

        /// <summary>
        /// 根据机种名自动补全产品别/客户别（仅填充空字段）：
        /// 产品别 = 机种名开始 2 位代码，客户别 = 机种名第 8 位起的 2 位代码；代码映射缺失时回退机种映射表。
        /// </summary>
        private void UpdateModelMapping()
        {
            string model = txt_model?.Text?.Trim();
            if (string.IsNullOrWhiteSpace(model))
            {
                return;
            }
            if (string.IsNullOrWhiteSpace(txt_product.Text))
            {
                string product = _admin.FindProductByModel(model);
                txt_product.Text = product ?? _admin.FindModelMapping(model)?.Product ?? "";
            }
            if (string.IsNullOrWhiteSpace(txt_customer.Text))
            {
                string customer = _admin.FindCustomerByModel(model);
                txt_customer.Text = customer ?? _admin.FindModelMapping(model)?.Customer ?? "";
            }
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
            if (_editTarget == null && dp_startDate.SelectedDate is DateTime start && string.IsNullOrWhiteSpace(txt_jobNo.Text))
            {
                txt_jobNo.Text = _excelService.GenerateJobNo(start, "QRT");
            }
            UpdateAutoPlan();
        }

        private void Txt_Model_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (txt_product == null)
            {
                return;
            }
            UpdateModelMapping();
        }

        private void Btn_Save_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateRequired(requireRemark: true))
            {
                return;
            }

            // 工作编号：允许手动指定，留空时仍自动生成 QRT{年月}{编号}
            string jobNo = ResolveJobNo();

            // 工作编号格式与唯一性校验
            string jobNoError = PlanValidation.ValidateJobNo(jobNo);
            if (jobNoError != null)
            {
                _ = MessageBox.Show(jobNoError, LanguageService.Get("Cap_FormatValidationFailed"));
                return;
            }
            long selfId = _editTarget?.Id ?? 0;
            if (_db.FreeSql.Select<Plan>().Where(p => p.JobNo == jobNo && p.Id != selfId).Any())
            {
                _ = MessageBox.Show(string.Format(LocalizationHelper.Get("Msg_JobNoExistsFormat"), jobNo), LanguageService.Get("Cap_Info"));
                return;
            }

            PlanResult = BuildPlanFromInputs(jobNo);

            // 非模态窗口：触发 Saved 事件后关闭，由调用方处理暂存/提审
            Saved?.Invoke(PlanResult, _editTarget?.Id ?? 0);
            Close();
        }

        /// <summary>
        /// 「转为领用」：把当前填写内容交给调用方，由它打开「领退表新增」并带上这些值
        /// （领退表那边有自己的必填项，这里只校验计划侧的必要信息，备注不强制）
        /// </summary>
        private void Btn_ConvertToRequisition_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateRequired(requireRemark: false))
            {
                return;
            }
            ConvertToRequisitionRequested?.Invoke(BuildPlanFromInputs(ResolveJobNo()));
            Close();
        }

        /// <summary>
        /// 必填校验（与标签上的 * 对应）：测试项目/开始时间/阶段/机种名，保存时备注也算必填
        /// </summary>
        private bool ValidateRequired(bool requireRemark)
        {
            if (cb_testItem.SelectedItem == null)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_SelectTestItem"), LanguageService.Get("Cap_Info"));
                return false;
            }
            if (dp_startDate.SelectedDate == null)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillStartTime"), LanguageService.Get("Cap_Info"));
                return false;
            }
            if (cb_stage.SelectedItem == null)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_SelectStage"), LanguageService.Get("Cap_Info"));
                return false;
            }
            if (string.IsNullOrWhiteSpace(txt_model.Text))
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillModelName"), LanguageService.Get("Cap_Info"));
                return false;
            }
            if (requireRemark && string.IsNullOrWhiteSpace(txt_remark.Text))
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillRemark"), LanguageService.Get("Cap_Info"));
                return false;
            }
            return true;
        }

        /// <summary>
        /// 工作编号：手动填了就用填的，留空则按开始日期自动生成 QRT{年月}{编号}
        /// </summary>
        private string ResolveJobNo()
            => string.IsNullOrWhiteSpace(txt_jobNo.Text)
                ? _excelService.GenerateJobNo(dp_startDate.SelectedDate ?? DateTime.Today, "QRT")
                : txt_jobNo.Text.Trim();

        /// <summary>
        /// 把界面上的内容组装成计划记录（新增建一条、编辑在副本上改，保留 Id 与创建信息）
        /// </summary>
        private Plan BuildPlanFromInputs(string jobNo)
        {
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
            plan.ReportStatus = cb_reportStatus.SelectedItem as string;
            plan.Remark = string.IsNullOrWhiteSpace(txt_remark.Text) ? null : txt_remark.Text.Trim();
            plan.UpdatedBy = _permission.CurrentUser;
            plan.UpdatedAt = DateTime.Now;
            return plan;
        }
        
        private void Btn_Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private static Plan ClonePlan(Plan source)
            => Newtonsoft.Json.JsonConvert.DeserializeObject<Plan>(
                Newtonsoft.Json.JsonConvert.SerializeObject(source));
    }
}
