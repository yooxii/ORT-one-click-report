using NLog;
using ORT一键报告.Models;
using ORT一键报告.Services;
using ORT一键报告.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// WindowRequisitionEdit.xaml 的交互逻辑：领退表新增/编辑。
    /// 必填：領用日期/領料單据號/機種名稱/領出數量/S-N/REV./Work Order；
    /// 自动补全：D/C、線別、回线RT工令（可选）、计划表同步信息（测试项目/开始时间/阶段/工作编号/样品数/产品别/客户别/试验时间/负责人/结束日期）。
    /// 本对话框只构造结果，不写数据库；由调用方决定暂存或提审。
    /// </summary>
    public partial class WindowRequisitionEdit : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;
        private readonly IPermissionService _permission;
        private readonly AdminService _admin;
        private readonly PlanExcelService _excelService;
        private readonly Requisition _editTarget;

        /// <summary>
        /// 编辑模式下找到的关联计划记录（按工令/机种+日期匹配）
        /// </summary>
        private Plan _associatedPlan;

        /// <summary>
        /// 上传方式选择的SN源文件路径
        /// </summary>
        private string _uploadedSnFile;

        /// <summary>
        /// 最近一次由程序生成的工作编号 / 回线RT工令：
        /// 改日期时若字段内容还是它（用户没手动改过）就跟着刷新，手动改过的不覆盖
        /// </summary>
        private string _autoJobNo;
        private string _autoReturnRt;

        /// <summary>
        /// 构造的领退记录结果（由调用方处理：暂存或提审）
        /// </summary>
        public Requisition RequisitionResult { get; private set; }

        /// <summary>
        /// 构造的计划记录结果（同步新增计划）
        /// </summary>
        public Plan PlanResult { get; private set; }

        /// <summary>
        /// 保存成功事件（非模态窗口用）：参数为 (领退结果, 计划结果, 编辑目标Id)。
        /// 调用方订阅后处理暂存/提审逻辑，窗口自身只负责构造结果并关闭。
        /// </summary>
        public event Action<Requisition, Plan, long> Saved;

        public WindowRequisitionEdit(DatabaseService db, IPermissionService permission, AdminService admin,
            PlanExcelService excelService, Requisition editTarget = null,
            string defaultTestItem = null, string defaultStage = null)
        {
            InitializeComponent();
            _db = db;
            _permission = permission;
            _admin = admin;
            _excelService = excelService;
            _editTarget = editTarget;

            Title = editTarget == null ? "领退表新增" : "领退表编辑";

            // 字典初始化
            cb_testItem.ItemsSource = _admin.GetTestItems().Select(t => t.Name).ToList();
            cb_stage.ItemsSource = _admin.GetStages().Select(s => s.Name).ToList();
            cb_reportStatus.ItemsSource = Models.ReportStatusKind.All.ToList();

            // 新增时默认预选调用方传入的测试项目/阶段（当前计划表里用得最多的一项）；
            // 编辑时不预选（由 LoadFromRequisition / LoadAssociatedPlan 带入原值），
            // 「转为领用」路径由随后的 PrefillFromPlan 覆盖（调用方那边传 (null, null) 不会走到这里）。
            if (editTarget == null)
            {
                SetCombo(cb_testItem, defaultTestItem);
                SetCombo(cb_stage, defaultStage);
            }

            // 回线RT工令：默认「无需回线」——不勾选时留空并禁用输入
            ApplyNeedReturnState();

            // 领用日期变化：开始时间跟着走；工作编号/回线RT工令还是自动生成的那个（没手动改过）就一起刷新
            dp_reqDate.SelectedDateChanged += (s, e) => OnReqDateChanged();

            if (editTarget != null)
            {
                LoadFromRequisition(editTarget);
                LoadAssociatedPlan(editTarget);
            }
        }

        /// <summary>
        /// 领用日期变化后的联动：开始时间默认跟随；工作编号（及需要回线时的回线RT工令）跟着日期刷新，
        /// 修复「日期填错更正后编号还是旧的」
        /// </summary>
        private void OnReqDateChanged()
        {
            if (dp_reqDate.SelectedDate is not DateTime reqDate)
            {
                return;
            }
            if (_editTarget == null && dp_startDate != null && dp_reqDate.SelectedDate != null)
            {
                // 新增时开始时间默认与领用日期一致（与原有行为一致）
                dp_startDate.SelectedDate = dp_reqDate.SelectedDate;
            }
            // 编号未手动改过 → 跟着新日期重新生成
            if (_autoJobNo != null && txt_jobNo != null
                && string.Equals(txt_jobNo.Text?.Trim(), _autoJobNo, StringComparison.Ordinal))
            {
                RegenerateJobNo(false);
            }
            if (_autoReturnRt != null && txt_returnRt != null
                && string.Equals(txt_returnRt.Text?.Trim(), _autoReturnRt, StringComparison.Ordinal))
            {
                RegenerateReturnRt(false);
            }
            UpdateAutoPlan();
        }

        /* ###############################  加载  ################################ */

        /// <summary>
        /// 用计划表草稿预填（计划表新增窗口的「转为领用」入口用）：
        /// 领退信息里的机种/领出数量、计划同步区的测试项目/阶段/开始时间/负责人等一并带过来；
        /// 工作编号不带过来（RT 号由本窗口按领用日期另行生成，也可手动改）。
        /// </summary>
        public void PrefillFromPlan(Plan plan)
        {
            if (plan == null)
            {
                return;
            }
            // 先给领用日期，让 UpdateAutoPlan 能按日期生成 RT 工作编号
            dp_reqDate.SelectedDate = plan.StartDate ?? DateTime.Today;
            // 从计划表「转为领用」过来的：计划表同步信息已经填好了，直接展开（展开=同步建立计划表记录）
            exp_planSync.IsExpanded = true;
            txt_model.Text = plan.ModelName ?? "";
            txt_outQty.Text = plan.SampleSize ?? "";
            SetCombo(cb_testItem, plan.TestItem);
            SetCombo(cb_stage, plan.Stage);
            SetCombo(cb_reportStatus, plan.ReportStatus);
            dp_startDate.SelectedDate = plan.StartDate;
            txt_sampleSize.Text = plan.SampleSize ?? "";
            txt_product.Text = plan.Product ?? "";
            txt_customer.Text = plan.Customer ?? "";
            txt_owner.Text = plan.Owner ?? "";
            txt_testPeriod.Text = plan.TestPeriod ?? "";
            dp_endDate.SelectedDate = plan.EndDate;
            // 机种/测试项目联动补齐产品别、客户别、负责人、结束日期等空字段
            UpdateAutoPlan();
        }

        /// <summary>
        /// 编辑模式：查找领退记录对应的计划表记录并载入计划同步区。
        /// 匹配规则：备注含回线RT工令 → 备注含 WorkOrder → 机种+领用日期同开始日期。
        /// </summary>
        private void LoadAssociatedPlan(Requisition req)
        {
            List<Plan> plans = _db.FreeSql.Select<Plan>().ToList();
            _associatedPlan =
                (!string.IsNullOrWhiteSpace(req.ReturnRtOrder)
                    ? plans.FirstOrDefault(p => p.Remark != null && p.Remark.Contains(req.ReturnRtOrder)) : null)
                ?? (!string.IsNullOrWhiteSpace(req.WorkOrder)
                    ? plans.FirstOrDefault(p => p.Remark != null && p.Remark.Contains(req.WorkOrder)) : null)
                ?? plans.FirstOrDefault(p => p.ModelName == req.ModelName && p.StartDate != null && req.RequisitionDate != null
                    && p.StartDate.Value.Date == req.RequisitionDate.Value.Date);

            if (_associatedPlan == null)
            {
                return;
            }
            // 找到关联计划：展开「计划表同步信息」（展开=同步更新该计划记录）
            exp_planSync.IsExpanded = true;
            SetCombo(cb_testItem, _associatedPlan.TestItem);
            SetCombo(cb_stage, _associatedPlan.Stage);
            SetCombo(cb_reportStatus, _associatedPlan.ReportStatus);
            dp_startDate.SelectedDate = _associatedPlan.StartDate;
            txt_jobNo.Text = _associatedPlan.JobNo;
            txt_sampleSize.Text = _associatedPlan.SampleSize;
            txt_product.Text = _associatedPlan.Product;
            txt_customer.Text = _associatedPlan.Customer;
            txt_owner.Text = _associatedPlan.Owner;
            txt_testPeriod.Text = _associatedPlan.TestPeriod;
            dp_endDate.SelectedDate = _associatedPlan.EndDate;
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

        private void LoadFromRequisition(Requisition req)
        {
            dp_reqDate.SelectedDate = req.RequisitionDate;
            txt_reqNo.Text = req.RequisitionNo;
            txt_model.Text = req.ModelName;
            txt_outQty.Text = req.OutQty;
            txt_rev.Text = req.Rev;
            txt_workOrder.Text = req.WorkOrder;
            txt_dc.Text = req.DC;
            txt_lineNo.Text = req.LineNo;
            // 回线RT工令：有值即视为「需要回线」（勾选并把值显示出来），没有就留空=无需回线
            chk_needReturn.IsChecked = !string.IsNullOrWhiteSpace(req.ReturnRtOrder);
            txt_returnRt.Text = req.ReturnRtOrder ?? "";
            ApplyNeedReturnState();
            if (!string.IsNullOrWhiteSpace(req.SnFilePath))
            {
                rb_snFile.IsChecked = true;
                _uploadedSnFile = _db.ResolveAttachmentPath(req.SnFilePath);
                txt_snFileName.Text = req.SnFilePath;
            }
            else
            {
                txt_sn.Text = req.SN;
            }
            AutoFillFromWorkOrder();
        }

        /* ###############################  自动补全  ################################ */

        /// <summary>
        /// 从 Work Order 自动补全 D/C（倒数第三位起的两位）与 線別（倒数第六位起的三位）；仅填充空字段，允许手动修改
        /// </summary>
        private void AutoFillFromWorkOrder()
        {
            string wo = txt_workOrder?.Text?.Trim();
            if (string.IsNullOrWhiteSpace(wo) || wo.Length < 6)
            {
                return;
            }
            // 線別：倒数第六位起的三位字符串
            if (string.IsNullOrWhiteSpace(txt_lineNo.Text))
            {
                txt_lineNo.Text = wo.Substring(wo.Length - 6, 3);
            }
            // D/C：倒数第三位起的两位
            if (string.IsNullOrWhiteSpace(txt_dc.Text) && wo.Length >= 3)
            {
                txt_dc.Text = wo.Substring(wo.Length - 3, 2);
            }
            UpdateAutoPlan();
        }

        /// <summary>
        /// 生成工作编号 RT{年月}{编号}，并记住这个自动生成的值（改日期时据此判断是否跟随刷新）；
        /// force=true 时（点标签）日期缺失会给出提示
        /// </summary>
        private void RegenerateJobNo(bool force)
        {
            if (dp_reqDate.SelectedDate is not DateTime dt)
            {
                if (force)
                {
                    _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillReqDate"), LanguageService.Get("Cap_Info"));
                }
                return;
            }
            txt_jobNo.Text = _excelService.GenerateJobNo(dt, "RT");
            _autoJobNo = txt_jobNo.Text.Trim();
        }

        /// <summary>
        /// 生成回线RT工令 RTAH{年月}{编号}：点「回线RT工令」标签时按领用日期生成（已取消自动生成）；
        /// 点标签即视为需要回线，会一并勾上「需要回线」
        /// </summary>
        private void RegenerateReturnRt(bool force)
        {
            if (dp_reqDate.SelectedDate is not DateTime dt)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillReqDate"), LanguageService.Get("Cap_Info"));
                return;
            }
            if (force && chk_needReturn.IsChecked != true)
            {
                chk_needReturn.IsChecked = true;
            }
            txt_returnRt.Text = _excelService.GenerateReturnRtOrder(dt);
            _autoReturnRt = txt_returnRt.Text.Trim();
        }

        /// <summary>
        /// 「需要回线」勾选状态落到回线RT工令输入框：不勾选=无需回线（留空并禁用）
        /// </summary>
        private void ApplyNeedReturnState()
        {
            if (txt_returnRt == null || chk_needReturn == null)
            {
                return;
            }
            bool need = chk_needReturn.IsChecked == true;
            txt_returnRt.IsEnabled = need;
            if (!need)
            {
                txt_returnRt.Text = "";
                _autoReturnRt = null;
            }
        }

        /// <summary>
        /// 刷新计划表同步信息：新增时生成工作编号/样品数等；编辑时仅做测试项目联动，不覆盖已有值（仅填充空字段）
        /// </summary>
        private void UpdateAutoPlan()
        {
            if (dp_reqDate.SelectedDate is not DateTime reqDate)
            {
                return;
            }
            if (_editTarget == null)
            {
                // 开始时间默认与领用日期一致；工作编号/样品数自动生成（仅填充空字段）
                if (dp_startDate.SelectedDate == null)
                {
                    dp_startDate.SelectedDate = dp_reqDate.SelectedDate;
                }
                if (string.IsNullOrWhiteSpace(txt_jobNo.Text))
                {
                    RegenerateJobNo(false);
                }
                if (string.IsNullOrWhiteSpace(txt_sampleSize.Text))
                {
                    txt_sampleSize.Text = txt_outQty?.Text?.Trim();
                }
            }

            // 产品别/客户别：根据机种名代码规则查询（产品别=前2位，客户别=第8位起2位）
            string model = txt_model?.Text?.Trim();
            if (!string.IsNullOrWhiteSpace(model))
            {
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
            // 负责人/试验时间/结束日期：根据测试项目查询
            string testItem = cb_testItem.SelectedItem as string;
            if (!string.IsNullOrWhiteSpace(testItem))
            {
                TestItemCatalog item = _admin.GetTestItems().FirstOrDefault(t => t.Name == testItem);
                if (item != null)
                {
                    if (string.IsNullOrWhiteSpace(txt_owner.Text))
                    {
                        txt_owner.Text = item.Owner;
                    }
                    if (string.IsNullOrWhiteSpace(txt_testPeriod.Text))
                    {
                        txt_testPeriod.Text = item.Period;
                    }
                    if (dp_endDate.SelectedDate == null && int.TryParse(item.Period, out int hours))
                    {
                        DateTime start = dp_startDate.SelectedDate ?? reqDate;
                        dp_endDate.SelectedDate = start.AddHours(hours);
                    }
                }
            }
        }

        /* ###############################  事件函数  ################################ */

        private void Txt_Model_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (txt_product == null)
            {
                return;
            }
            UpdateAutoPlan();
        }

        private void SnMode_Changed(object sender, RoutedEventArgs e)
        {
            if (txt_sn == null || btn_snFile == null)
            {
                return;
            }
            bool isInput = rb_snInput.IsChecked == true;
            txt_sn.Visibility = isInput ? Visibility.Visible : Visibility.Collapsed;
            btn_snFile.Visibility = isInput ? Visibility.Collapsed : Visibility.Visible;
            txt_snFileName.Visibility = isInput ? Visibility.Collapsed : Visibility.Visible;
        }

        private void Btn_SnFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dialog = new()
            {
                Title = LanguageService.Get("Title_SelectSNFile"),
                Filter = "Excel文件|*.xls;*.xlsx;*.xlsm|文本文件|*.txt;*.csv|所有文件|*.*"
            };
            if (dialog.ShowDialog() == true)
            {
                _uploadedSnFile = dialog.FileName;
                txt_snFileName.Text = _uploadedSnFile;
            }
        }

        /// <summary>
        /// 「需要回线」勾选变化：勾上就允许填写回线RT工令（并聚焦），取消就清空表示无需回线
        /// </summary>
        private void Chk_NeedReturn_Changed(object sender, RoutedEventArgs e)
        {
            ApplyNeedReturnState();
            if (chk_needReturn.IsChecked == true && string.IsNullOrWhiteSpace(txt_returnRt.Text))
            {
                txt_returnRt.Focus();
            }
        }

        /// <summary>
        /// 点「工作编号」标签：按当前领用日期重新生成（日期填错更正后用它刷新）
        /// </summary>
        private void Lbl_JobNo_Click(object sender, System.Windows.Input.MouseButtonEventArgs e) => RegenerateJobNo(true);

        /// <summary>
        /// 点「回线RT工令」标签：按当前领用日期生成回线RT工令（不再自动生成）
        /// </summary>
        private void Lbl_ReturnRt_Click(object sender, System.Windows.Input.MouseButtonEventArgs e) => RegenerateReturnRt(true);

        /// <summary>
        /// 「计划表同步信息」展开时补齐计划侧能自动带的字段（只填空字段，不覆盖已填内容）
        /// </summary>
        private void Exp_PlanSync_Toggled(object sender, RoutedEventArgs e)
        {
            if (exp_planSync?.IsExpanded == true)
            {
                UpdateAutoPlan();
            }
        }

        private void Cb_TestItem_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (txt_owner == null)
            {
                return;
            }
            UpdateAutoPlan();
        }

        private void Btn_Save_Click(object sender, RoutedEventArgs e)
        {
            // 必填校验
            if (dp_reqDate.SelectedDate == null)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillReqDate"), LanguageService.Get("Cap_Info"));
                return;
            }
            if (string.IsNullOrWhiteSpace(txt_reqNo.Text))
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillDocNo"), LanguageService.Get("Cap_Info"));
                return;
            }
            if (string.IsNullOrWhiteSpace(txt_model.Text))
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillModelName"), LanguageService.Get("Cap_Info"));
                return;
            }
            if (string.IsNullOrWhiteSpace(txt_outQty.Text))
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillQty"), LanguageService.Get("Cap_Info"));
                return;
            }
            if (string.IsNullOrWhiteSpace(txt_rev.Text))
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillRev"), LanguageService.Get("Cap_Info"));
                return;
            }
            if (string.IsNullOrWhiteSpace(txt_workOrder.Text))
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillWorkOrder"), LanguageService.Get("Cap_Info"));
                return;
            }
            if (rb_snInput.IsChecked == true && string.IsNullOrWhiteSpace(txt_sn.Text))
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_FillSN"), LanguageService.Get("Cap_Info"));
                return;
            }
            if (rb_snFile.IsChecked == true && _uploadedSnFile == null)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_SelectSNFile"), LanguageService.Get("Cap_Info"));
                return;
            }
            // 自定义序列号：提交时按「每行一个」检查重复，有重复先问用户要不要重新输入
            if (rb_snInput.IsChecked == true && !ConfirmDuplicateSn())
            {
                txt_sn.Focus();
                return;
            }
            // 回线：勾了「需要回线」就必须有回线RT工令；不勾选则留空表示无需回线
            if (chk_needReturn.IsChecked == true && string.IsNullOrWhiteSpace(txt_returnRt.Text))
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_NeedReturnOrder"), LanguageService.Get("Cap_Info"));
                return;
            }
            // 计划表同步信息：展开=同步（建立/更新计划表记录，下列字段必填），折叠=只登记领退信息
            bool syncPlan = exp_planSync.IsExpanded;
            string jobNo = null;
            if (syncPlan)
            {
                if (_editTarget != null && _associatedPlan == null)
                {
                    _ = MessageBox.Show(LocalizationHelper.Get("Msg_NoAssociatedPlan"), LanguageService.Get("Cap_Info"));
                    return;
                }
                if (cb_testItem.SelectedItem == null)
                {
                    _ = MessageBox.Show(LocalizationHelper.Get("Msg_SelectTestItem"), LanguageService.Get("Cap_Info"));
                    return;
                }
                if (cb_stage.SelectedItem == null)
                {
                    _ = MessageBox.Show(LocalizationHelper.Get("Msg_SelectStage"), LanguageService.Get("Cap_Info"));
                    return;
                }
                // 工作编号：允许手动指定，留空时按领用日期自动生成 RT{年月}{编号}
                jobNo = string.IsNullOrWhiteSpace(txt_jobNo.Text)
                    ? _excelService.GenerateJobNo(dp_reqDate.SelectedDate ?? DateTime.Today, "RT")
                    : txt_jobNo.Text.Trim();
                string jobNoError = PlanValidation.ValidateJobNo(jobNo);
                if (jobNoError != null)
                {
                    _ = MessageBox.Show(jobNoError, LanguageService.Get("Cap_FormatValidationFailed"));
                    return;
                }
                long planId = _associatedPlan?.Id ?? 0;
                if (_db.FreeSql.Select<Plan>().Where(p => p.JobNo == jobNo && p.Id != planId).Any())
                {
                    _ = MessageBox.Show(string.Format(LocalizationHelper.Get("Msg_JobNoExistsFormat"), jobNo),
                        LanguageService.Get("Cap_Info"));
                    return;
                }
            }

            // 唯一键校验
            long selfId = _editTarget?.Id ?? 0;
            if (_db.FreeSql.Select<Requisition>().Where(r => r.RequisitionNo == txt_reqNo.Text.Trim() && r.Id != selfId).Any())
            {
                _ = MessageBox.Show($"領料單据號 [{txt_reqNo.Text.Trim()}] 已存在", LanguageService.Get("Cap_Info"));
                return;
            }

            // 构造领退记录
            Requisition req = _editTarget == null
                ? new Requisition { CreatedBy = _permission.CurrentUser, CreatedAt = DateTime.Now }
                : CloneReq(_editTarget);
            req.RequisitionDate = dp_reqDate.SelectedDate;
            req.RequisitionNo = txt_reqNo.Text.Trim();
            req.ModelName = txt_model.Text.Trim();
            req.OutQty = txt_outQty.Text.Trim();
            req.Rev = txt_rev.Text.Trim();
            req.WorkOrder = txt_workOrder.Text.Trim();
            req.DC = txt_dc.Text.Trim();
            req.LineNo = txt_lineNo.Text.Trim();
            req.ReturnRtOrder = string.IsNullOrWhiteSpace(txt_returnRt.Text) ? null : txt_returnRt.Text.Trim();
            req.UpdatedBy = _permission.CurrentUser;
            req.UpdatedAt = DateTime.Now;

            // S/N
            if (rb_snInput.IsChecked == true)
            {
                req.SN = txt_sn.Text.Trim();
                req.SnFilePath = null;
            }
            else
            {
                string existingFile = _editTarget == null ? null : _db.ResolveAttachmentPath(_editTarget.SnFilePath);
                if (_editTarget != null && string.Equals(_uploadedSnFile, existingFile, StringComparison.OrdinalIgnoreCase))
                {
                    req.SnFilePath = _editTarget.SnFilePath;
                }
                else
                {
                    string savedName = SaveSnFile(_uploadedSnFile, req.RequisitionNo, req.ModelName);
                    if (savedName == null)
                    {
                        return;
                    }
                    req.SnFilePath = savedName;
                }
            }

            RequisitionResult = req;

            // 计划表同步：展开才构造/更新计划记录（新增建立 / 编辑更新关联计划）；折叠时 PlanResult=null，
            // 调用方只暂存领退记录，不再顺带建立计划表记录
            if (!syncPlan)
            {
                PlanResult = null;
            }
            else
            {
                Plan plan = _associatedPlan != null
                    ? ClonePlan(_associatedPlan)   // 编辑关联计划：保持 Id/创建信息
                    : new Plan { CreatedBy = _permission.CurrentUser, CreatedAt = DateTime.Now };
                plan.JobNo = jobNo;
                plan.TestItem = cb_testItem.SelectedItem as string;
                plan.StartDate = dp_startDate.SelectedDate ?? dp_reqDate.SelectedDate;
                plan.Stage = cb_stage.SelectedItem as string;
                // 样品数留空时：新增按领出数量兜底；编辑保持原值（不把原有空值改成领出数量）
                plan.SampleSize = string.IsNullOrWhiteSpace(txt_sampleSize.Text)
                    ? (_editTarget == null ? req.OutQty : null)
                    : txt_sampleSize.Text.Trim();
                plan.ModelName = req.ModelName;
                plan.Product = string.IsNullOrWhiteSpace(txt_product.Text) ? null : txt_product.Text.Trim();
                plan.Customer = string.IsNullOrWhiteSpace(txt_customer.Text) ? null : txt_customer.Text.Trim();
                plan.Owner = string.IsNullOrWhiteSpace(txt_owner.Text) ? null : txt_owner.Text.Trim();
                plan.TestPeriod = string.IsNullOrWhiteSpace(txt_testPeriod.Text) ? null : txt_testPeriod.Text.Trim();
                plan.EndDate = dp_endDate.SelectedDate;
                plan.Status = plan.Status ?? "Ongoing";
                plan.ReportStatus = cb_reportStatus.SelectedItem as string;
                plan.UpdatedBy = _permission.CurrentUser;
                plan.UpdatedAt = DateTime.Now;
                PlanResult = plan;
            }

            // 非模态窗口：触发 Saved 事件后关闭，由调用方处理暂存/提审
            Saved?.Invoke(RequisitionResult, PlanResult, _editTarget?.Id ?? 0);
            Close();
        }
        
        private void Btn_Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        /* ###############################  序列号重复检查  ################################ */

        /// <summary>
        /// 自定义序列号提交前的重复检查：按「每行一个序列号」拆分，
        /// 既查本次输入内部的重号，也查领退记录里已有的序列号（编辑时排除本记录）。
        /// 有重复时弹窗询问，返回 false 表示用户选择「重新输入」（不保存）。
        /// </summary>
        private bool ConfirmDuplicateSn()
        {
            List<string> problems = FindDuplicateSnProblems();
            if (problems.Count == 0)
            {
                return true;
            }
            MessageBoxResult choice = MessageBox.Show(
                string.Format(LocalizationHelper.Get("Msg_SnDuplicateFormat"), string.Join("；", problems)),
                LanguageService.Get("Cap_Info"), MessageBoxButton.YesNo, MessageBoxImage.Warning);
            return choice != MessageBoxResult.Yes;   // 「是」= 重新输入（放弃本次保存）
        }

        /// <summary>
        /// 找出重复的序列号（描述文本列表）：本次输入内的重号 + 领退记录里已有的序列号。
        /// 空列表表示没有重复。
        /// </summary>
        private List<string> FindDuplicateSnProblems()
        {
            List<string> sns = ParseSnLines(txt_sn.Text);
            List<string> problems = [];
            if (sns.Count == 0)
            {
                return problems;
            }

            foreach (string sn in sns.GroupBy(s => s, StringComparer.OrdinalIgnoreCase)
                                     .Where(g => g.Count() > 1)
                                     .Select(g => g.Key))
            {
                problems.Add($"{sn}（{LocalizationHelper.Get("Msg_SnDuplicateInInput")}）");
            }

            // 已有领退记录里的序列号（序列号 → 领料单据号）
            Dictionary<string, string> existing = new(StringComparer.OrdinalIgnoreCase);
            long selfId = _editTarget?.Id ?? 0;
            foreach (Requisition other in _db.FreeSql.Select<Requisition>().Where(r => r.Id != selfId && r.SN != null).ToList())
            {
                foreach (string sn in ParseSnLines(other.SN))
                {
                    if (!existing.ContainsKey(sn))
                    {
                        existing[sn] = other.RequisitionNo ?? "";
                    }
                }
            }
            foreach (string sn in sns.Where(existing.ContainsKey).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                problems.Add($"{sn}（{string.Format(LocalizationHelper.Get("Msg_SnDuplicateInDbFormat"), existing[sn])}）");
            }
            return problems;
        }

        /// <summary>
        /// 把自定义序列号输入拆成一行一个（去空行与首尾空白）
        /// </summary>
        private static List<string> ParseSnLines(string snText)
        {
            if (string.IsNullOrWhiteSpace(snText))
            {
                return [];
            }
            return snText.Split(['\n', '\r'], StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim())
                .Where(s => s.Length > 0)
                .ToList();
        }

        private string SaveSnFile(string sourcePath, string key, string modelName)
        {
            try
            {
                string name = $"{DateTime.Now:MMdd}_{Clean(key ?? "无编号")}_{Clean(modelName ?? "无机种名")}_{Clean(Path.GetFileName(sourcePath))}";
                string fullPath = Path.Combine(_db.OleDir, name);
                if (File.Exists(fullPath))
                {
                    name = $"{DateTime.Now:MMddHHmmss}_{Clean(key ?? "无编号")}_{Clean(modelName ?? "无机种名")}_{Clean(Path.GetFileName(sourcePath))}";
                    fullPath = Path.Combine(_db.OleDir, name);
                }
                File.Copy(sourcePath, fullPath, true);
                return name;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "保存上传的SN文件失败");
                _ = MessageBox.Show($"保存上传文件失败:\n{ex.Message}", LanguageService.Get("Cap_Error"));
                return null;
            }
        }

        private static Requisition CloneReq(Requisition source)
            => Newtonsoft.Json.JsonConvert.DeserializeObject<Requisition>(
                Newtonsoft.Json.JsonConvert.SerializeObject(source));

        private static Plan ClonePlan(Plan source)
            => Newtonsoft.Json.JsonConvert.DeserializeObject<Plan>(
                Newtonsoft.Json.JsonConvert.SerializeObject(source));

        private static string Clean(string name)
        {
            string cleaned = Regex.Replace(name ?? "", $"[{Regex.Escape(new string(Path.GetInvalidFileNameChars()))}]", "_").Trim();
            return cleaned == "" ? "_" : cleaned;
        }
    }
}
