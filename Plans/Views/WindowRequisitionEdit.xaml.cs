using NLog;
using ORT一键报告.Models;
using ORT一键报告.Plans.ViewModels;
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
    /// 必填：領用日期/領料單据號/機種名稱/領出數量/單體去向（入库/报废）/S-N/REV./Work Order；
    /// 自动补全：D/C、線別、回线RT工令（单体去向=入库时必填）、计划表同步信息（测试项目/开始时间/阶段/工作编号/样品数/产品别/客户别/试验时间/负责人/结束日期）。
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
        /// 最近一次由完工令（Work Order）自动补全的 D/C 与线别：
        /// 改完工令时若字段内容还是它（用户没手动改过）就跟着刷新
        /// </summary>
        private string _autoDc;
        private string _autoLineNo;

        /// <summary>
        /// 本窗口已保存（含「保存并继续」的前几条）的工作编号 / 回线RT工令：
        /// 这些还只在内存里、库中查不到，重新生成时要跳过，避免连续录入重号
        /// </summary>
        private readonly HashSet<string> _issuedJobNos = new(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _issuedReturnRt = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// 本窗口「保存并继续」刚存过的「机种 → 版本」：提示优先用它，其次才是库里的历史记录
        /// </summary>
        private readonly Dictionary<string, string> _sessionRevByModel = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// 当前机种最近一次使用过的版本（版本输入框留空时保存就用它），没有历史时为 null
        /// </summary>
        private string _suggestedRev;

        /// <summary>
        /// 版本提示对应的机种文本（机种没变就不重复查库）
        /// </summary>
        private string _revHintModel;

        /// <summary>
        /// 程序化改字段（加载/重置）期间置位：抑制机种、完工令等联动的自动判定，避免和随后赋的原值打架
        /// </summary>
        private bool _suppressAuto;

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
            cb_disposition.ItemsSource = Models.RequisitionDispositionKind.All.ToList();

            // 新增时默认预选调用方传入的测试项目/阶段（当前计划表里用得最多的一项）；
            // 编辑时不预选（由 LoadFromRequisition / LoadAssociatedPlan 带入原值），
            // 「转为领用」路径由随后的 PrefillFromPlan 覆盖（调用方那边传 (null, null) 不会走到这里）。
            if (editTarget == null)
            {
                SetCombo(cb_testItem, defaultTestItem);
                SetCombo(cb_stage, defaultStage);
            }

            // 回线RT工令：由「单体去向」驱动——入库才需要（启用且保存时必填），报废禁用并留空
            cb_disposition.SelectionChanged += (s, e) => ApplyDispositionState();
            ApplyDispositionState();

            // 「保存并继续」只在新增时有意义（编辑模式保存后表单没有下一条可录）
            btn_saveAndContinue.Visibility = editTarget == null ? Visibility.Visible : Visibility.Collapsed;

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
            // 回填期间抑制联动：机种/完工令的自动判定不能盖掉库里存的原值（回填完再单独补空字段）
            _suppressAuto = true;
            try
            {
                dp_reqDate.SelectedDate = req.RequisitionDate;
                txt_reqNo.Text = req.RequisitionNo;
                txt_model.Text = req.ModelName;
                txt_outQty.Text = req.OutQty;
                txt_rev.Text = req.Rev;
                txt_workOrder.Text = req.WorkOrder;
                txt_dc.Text = req.DC;
                txt_lineNo.Text = req.LineNo;
                // 单体去向：入库/报废（旧记录可能为空，保存时按必填校验拦截）
                SetCombo(cb_disposition, req.Disposition);
                // 回线RT工令：按单体去向决定是否启用（报废时禁用并清空，保存即归位）
                txt_returnRt.Text = req.ReturnRtOrder ?? "";
                ApplyDispositionState();
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
            }
            finally
            {
                _suppressAuto = false;
            }
            UpdateRevHint();
            AutoFillFromWorkOrder();
        }

        /* ###############################  自动补全  ################################ */

        /// <summary>
        /// 从 Work Order 自动补全 D/C（倒数第三位起的两位）与 線別（倒数第六位起的三位）：
        /// 只填空字段，以及「上一次自动带出来、用户没手改过」的值（改工令会跟着刷新）；用户手改过的不覆盖。
        /// </summary>
        private void AutoFillFromWorkOrder()
        {
            if (_suppressAuto || txt_workOrder == null || txt_lineNo == null || txt_dc == null)
            {
                return;
            }
            string wo = txt_workOrder.Text?.Trim();
            if (string.IsNullOrWhiteSpace(wo) || wo.Length < 6)
            {
                // 工令清空或太短：把上一次自动带出的值一并清掉（用户手改过的不动）
                ClearAutoFilled(txt_lineNo, ref _autoLineNo);
                ClearAutoFilled(txt_dc, ref _autoDc);
                return;
            }
            // 線別：倒数第六位起的三位字符串
            string lineNo = RequisitionEditRules.LineNoFromWorkOrder(wo);
            if (lineNo != null && CanOverwrite(txt_lineNo.Text, _autoLineNo))
            {
                txt_lineNo.Text = lineNo;
                _autoLineNo = lineNo;
            }
            // D/C：倒数第三位起的两位
            string dc = RequisitionEditRules.DcFromWorkOrder(wo);
            if (dc != null && CanOverwrite(txt_dc.Text, _autoDc))
            {
                txt_dc.Text = dc;
                _autoDc = dc;
            }
            UpdateAutoPlan();
        }

        /// <summary>
        /// 能否覆盖：字段为空，或内容仍是上一次自动补全的值（用户没手改过）
        /// </summary>
        private static bool CanOverwrite(string current, string autoValue)
            => string.IsNullOrWhiteSpace(current)
               || (autoValue != null && string.Equals(current.Trim(), autoValue, StringComparison.Ordinal));

        /// <summary>
        /// 清掉上一次自动补全写入的值（内容已被用户改过则保留）
        /// </summary>
        private static void ClearAutoFilled(System.Windows.Controls.TextBox box, ref string autoValue)
        {
            if (autoValue != null && string.Equals(box.Text?.Trim(), autoValue, StringComparison.Ordinal))
            {
                box.Text = "";
            }
            autoValue = null;
        }

        /// <summary>
        /// 「版本」输入框的最近版本提示：按当前机种查最近一次用过的版本（本窗口刚存过的优先），
        /// 显示为输入框里的水印文字；用户不填时保存就用它
        /// </summary>
        private void UpdateRevHint()
        {
            string model = txt_model?.Text?.Trim();
            if (model != _revHintModel)
            {
                _revHintModel = model;
                _suggestedRev = null;
                if (!string.IsNullOrWhiteSpace(model))
                {
                    // 本窗口「保存并继续」刚存过的版本最"近"，优先于库里的历史
                    _ = _sessionRevByModel.TryGetValue(model, out _suggestedRev);
                    if (_suggestedRev == null)
                    {
                        long selfId = _editTarget?.Id ?? 0;
                        List<Requisition> rows = _db.FreeSql.Select<Requisition>()
                            .Where(r => r.ModelName == model && r.Id != selfId)
                            .OrderByDescending(r => r.RequisitionDate)
                            .OrderByDescending(r => r.Id)
                            .ToList();
                        _suggestedRev = rows.Select(r => r.Rev?.Trim())
                            .FirstOrDefault(rev => !string.IsNullOrWhiteSpace(rev));
                    }
                }
            }
            RefreshRevHintVisual();
        }

        /// <summary>
        /// 刷新版本提示水印：输入框为空且有历史版本时显示（有内容就藏起来）
        /// </summary>
        private void RefreshRevHintVisual()
        {
            if (txt_revHint == null || txt_rev == null)
            {
                return;
            }
            string hint = _suggestedRev == null
                ? null
                : string.Format(LocalizationHelper.Get("ReqEdit_LastRevHintFormat"), _suggestedRev);
            bool show = hint != null && string.IsNullOrWhiteSpace(txt_rev.Text);
            txt_revHint.Text = show ? hint : "";
            txt_revHint.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
            txt_rev.ToolTip = hint;
        }

        /// <summary>
        /// 机种联动「单体去向」：机种名称以 W 开头 → 报废，其余 → 入库（用户仍可手动改）
        /// </summary>
        private void ApplyDispositionFromModel()
        {
            if (_suppressAuto)
            {
                return;
            }
            string model = txt_model?.Text?.Trim();
            if (string.IsNullOrWhiteSpace(model) || cb_disposition == null)
            {
                return;
            }
            SetCombo(cb_disposition, RequisitionEditRules.DispositionFromModel(model));
        }

        /// <summary>
        /// 生成工作编号 RT{年月}{编号}（跳过本窗口已经发出去的号），
        /// 并记住这个自动生成的值（改日期时据此判断是否跟随刷新）
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
            txt_jobNo.Text = GenerateUniqueJobNo(dt);
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
            if (cb_disposition.SelectedItem as string != RequisitionDispositionKind.StockIn)
            {
                if (force)
                {
                    _ = MessageBox.Show(LocalizationHelper.Get("Msg_ReturnRtNeedsStockIn"), LanguageService.Get("Cap_Info"));
                }
                return;
            }
            txt_returnRt.Text = GenerateUniqueReturnRt(dt);
            _autoReturnRt = txt_returnRt.Text.Trim();
        }

        /// <summary>
        /// 按当月已有编号生成工作编号；与本窗口本次会话已发出的编号重号时往后顺延
        /// （「保存并继续」连续录入时这些号还没落库，否则会连出两个一样的号）
        /// </summary>
        private string GenerateUniqueJobNo(DateTime date)
        {
            string jobNo = _excelService.GenerateJobNo(date, "RT");
            while (_issuedJobNos.Contains(jobNo))
            {
                string next = RequisitionEditRules.NextJobNo(jobNo);
                if (next == null)
                {
                    break;
                }
                jobNo = next;
            }
            return jobNo;
        }

        /// <summary>
        /// 按当月已有回线工令生成编号；与本窗口本次会话已发出的编号重号时往后顺延
        /// </summary>
        private string GenerateUniqueReturnRt(DateTime date)
        {
            string returnRt = _excelService.GenerateReturnRtOrder(date);
            while (_issuedReturnRt.Contains(returnRt))
            {
                string next = RequisitionEditRules.NextReturnRtOrder(returnRt);
                if (next == null)
                {
                    break;
                }
                returnRt = next;
            }
            return returnRt;
        }

        /// <summary>
        /// 单体去向驱动回线RT工令输入框：入库=需要回线（启用、保存时必填）；
        /// 报废=无需回线（禁用并清空，避免留下无意义工令）
        /// </summary>
        private void ApplyDispositionState()
        {
            if (txt_returnRt == null || cb_disposition == null)
            {
                return;
            }
            bool needReturn = cb_disposition.SelectedItem as string == RequisitionDispositionKind.StockIn;
            txt_returnRt.IsEnabled = needReturn;
            if (!needReturn)
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

        /// <summary>
        /// 机种变化：带出产品别/客户别、按机种判定单体去向，并刷新「最近版本」提示
        /// </summary>
        private void Txt_Model_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (txt_product == null)
            {
                return;
            }
            ApplyDispositionFromModel();
            UpdateRevHint();
            UpdateAutoPlan();
        }

        /// <summary>
        /// 领出数量变化：样品数还是空的就跟着填（判单体去向只看机种，不在这里重复判定）
        /// </summary>
        private void Txt_OutQty_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            if (txt_product == null)
            {
                return;
            }
            UpdateAutoPlan();
        }

        /// <summary>
        /// Work Order 变化：立即补全 D/C（界面上显示为周期）与线别
        /// </summary>
        private void Txt_WorkOrder_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
            => AutoFillFromWorkOrder();

        /// <summary>
        /// 版本输入框变化：有内容时隐藏「最近版本」水印提示
        /// </summary>
        private void Txt_Rev_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
            => RefreshRevHintVisual();

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

        private void Btn_Save_Click(object sender, RoutedEventArgs e) => SaveOnce(false);

        /// <summary>
        /// 「保存并继续」：先像「保存」一样存下当前这条，再把表单清空留在窗口里录下一条
        /// </summary>
        private void Btn_SaveAndContinue_Click(object sender, RoutedEventArgs e) => SaveOnce(true);

        /// <summary>
        /// 保存当前这条领退（及同步的计划表记录）：continueAfterSave=true 时保存后重置表单留在窗口，
        /// false 时触发 Saved 事件后关闭窗口。本方法只构造结果，写库/提审由调用方处理。
        /// </summary>
        private void SaveOnce(bool continueAfterSave)
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
            // 版本：留空时用「最近一次该机种的版本」提示值（输入框里有水印），提示也没有才要求填写
            string rev = txt_rev.Text?.Trim();
            if (string.IsNullOrWhiteSpace(rev))
            {
                rev = _suggestedRev;
            }
            if (string.IsNullOrWhiteSpace(rev))
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
            // 回线：单体去向为入库时必须填回线RT工令；报废无需回线（输入框已禁用清空）
            if (cb_disposition.SelectedItem as string == RequisitionDispositionKind.StockIn
                && string.IsNullOrWhiteSpace(txt_returnRt.Text))
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_NeedReturnOrder"), LanguageService.Get("Cap_Info"));
                return;
            }
            // 回线RT工令：格式（RTAH + 4位年月 + 至少2位编号，可以为空）与重号校验
            string returnRt = string.IsNullOrWhiteSpace(txt_returnRt.Text) ? null : txt_returnRt.Text.Trim();
            string returnRtError = PlanValidation.ValidateReturnRtOrder(returnRt);
            if (returnRtError != null)
            {
                _ = MessageBox.Show(returnRtError, LanguageService.Get("Cap_FormatValidationFailed"));
                txt_returnRt.Focus();
                return;
            }
            if (returnRt != null && (_issuedReturnRt.Contains(returnRt) || IsReturnRtDuplicated(returnRt)))
            {
                _ = MessageBox.Show(string.Format(LocalizationHelper.Get("Msg_ReturnRtExistsFormat"), returnRt),
                    LanguageService.Get("Cap_Info"));
                txt_returnRt.Focus();
                return;
            }
            // 单体去向（入库/报废二选一）：新增与编辑都必须选择
            if (cb_disposition.SelectedItem == null)
            {
                _ = MessageBox.Show(LocalizationHelper.Get("Msg_SelectDisposition"), LanguageService.Get("Cap_Info"));
                cb_disposition.Focus();
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
                    ? GenerateUniqueJobNo(dp_reqDate.SelectedDate ?? DateTime.Today)
                    : txt_jobNo.Text.Trim();
                string jobNoError = PlanValidation.ValidateJobNo(jobNo);
                if (jobNoError != null)
                {
                    _ = MessageBox.Show(jobNoError, LanguageService.Get("Cap_FormatValidationFailed"));
                    return;
                }
                long planId = _associatedPlan?.Id ?? 0;
                if (_issuedJobNos.Contains(jobNo)
                    || _db.FreeSql.Select<Plan>().Where(p => p.JobNo == jobNo && p.Id != planId).Any())
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
            req.Rev = rev;
            req.WorkOrder = txt_workOrder.Text.Trim();
            req.DC = txt_dc.Text.Trim();
            req.LineNo = txt_lineNo.Text.Trim();
            req.Disposition = cb_disposition.SelectedItem as string;
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
                // 备注：默认写「工令」（Work Order），已有备注的关联计划不覆盖
                if (string.IsNullOrWhiteSpace(plan.Remark))
                {
                    plan.Remark = req.WorkOrder;
                }
                plan.UpdatedBy = _permission.CurrentUser;
                plan.UpdatedAt = DateTime.Now;
                PlanResult = plan;
            }

            // 记下这条已经发出去的编号/版本：这些还只在内存里（库中查不到），
            // 「保存并继续」的下一条要据此避重，版本提示也优先用它
            if (!string.IsNullOrWhiteSpace(jobNo))
            {
                _ = _issuedJobNos.Add(jobNo);
            }
            if (!string.IsNullOrWhiteSpace(returnRt))
            {
                _ = _issuedReturnRt.Add(returnRt);
            }
            if (!string.IsNullOrWhiteSpace(req.ModelName) && !string.IsNullOrWhiteSpace(req.Rev))
            {
                _sessionRevByModel[req.ModelName.Trim()] = req.Rev.Trim();
            }

            // 非模态窗口：触发 Saved 事件，由调用方处理暂存/提审；
            // 「保存并继续」把表单清空留在窗口里录下一条，「保存」则关闭
            Saved?.Invoke(RequisitionResult, PlanResult, _editTarget?.Id ?? 0);
            if (continueAfterSave)
            {
                ResetForNextEntry();
            }
            else
            {
                Close();
            }
        }

        /// <summary>
        /// 「保存并继续」后的表单重置：清空上一条的内容（保留领用日期与计划同步区的选择/展开状态），
        /// 并重新按当前日期生成工作编号（跳过刚存的那条，避免重号）
        /// </summary>
        private void ResetForNextEntry()
        {
            _suppressAuto = true;
            try
            {
                _associatedPlan = null;
                _autoJobNo = null;
                _autoReturnRt = null;
                _autoDc = null;
                _autoLineNo = null;
                _uploadedSnFile = null;

                txt_reqNo.Text = "";
                txt_model.Text = "";
                txt_outQty.Text = "";
                txt_rev.Text = "";
                txt_workOrder.Text = "";
                txt_dc.Text = "";
                txt_lineNo.Text = "";
                txt_returnRt.Text = "";
                txt_sn.Text = "";
                txt_snFileName.Text = "";
                rb_snInput.IsChecked = true;
                // 单体去向回到未选，等下一个机种进来再按规则判定
                cb_disposition.SelectedItem = null;
                txt_jobNo.Text = "";
                txt_sampleSize.Text = "";
                txt_product.Text = "";
                txt_customer.Text = "";
                // 保留：领用日期、计划同步区（测试项目/阶段/报告状态/负责人/试验时间/结束日期）与展开状态
                _revHintModel = null;
                UpdateRevHint();
            }
            finally
            {
                _suppressAuto = false;
            }
            // 按保留的日期补齐开始时间/工作编号等
            UpdateAutoPlan();
            txt_reqNo.Focus();
        }
        
        private void Btn_Cancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        /// <summary>
        /// 回线RT工令是否已被其他领退记录占用（不区分大小写；编辑时排除本记录）
        /// </summary>
        private bool IsReturnRtDuplicated(string returnRt)
        {
            long selfId = _editTarget?.Id ?? 0;
            return _db.FreeSql.Select<Requisition>()
                .Where(r => r.Id != selfId && r.ReturnRtOrder != null)
                .ToList(r => r.ReturnRtOrder)
                .Any(x => string.Equals(x?.Trim(), returnRt, StringComparison.OrdinalIgnoreCase));
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
