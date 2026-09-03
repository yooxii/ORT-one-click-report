using NLog;
using ORT一键报告.Models;
using ORT一键报告.Services;
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
        /// 构造的领退记录结果（由调用方处理：暂存或提审）
        /// </summary>
        public Requisition RequisitionResult { get; private set; }

        /// <summary>
        /// 构造的计划记录结果（同步新增计划）
        /// </summary>
        public Plan PlanResult { get; private set; }

        public WindowRequisitionEdit(DatabaseService db, IPermissionService permission, AdminService admin,
            PlanExcelService excelService, Requisition editTarget = null)
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

            // 开始时间默认与领用日期联动
            dp_reqDate.SelectedDateChanged += (s, e) =>
            {
                if (dp_startDate != null && dp_reqDate.SelectedDate != null && _editTarget == null)
                {
                    dp_startDate.SelectedDate = dp_reqDate.SelectedDate;
                    UpdateAutoPlan();
                }
            };

            if (editTarget != null)
            {
                LoadFromRequisition(editTarget);
                LoadAssociatedPlan(editTarget);
            }
        }

        /* ###############################  加载  ################################ */

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
            SetCombo(cb_testItem, _associatedPlan.TestItem);
            SetCombo(cb_stage, _associatedPlan.Stage);
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
            if (req.ReturnRtOrder != null)
            {
                chk_genReturnRt.IsChecked = false;
                txt_returnRt.Text = req.ReturnRtOrder;
                txt_returnRt.IsReadOnly = false;
            }
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
        /// 刷新回线RT工令（自动生成时）
        /// </summary>
        private void UpdateReturnRt()
        {
            if (chk_genReturnRt.IsChecked == true && dp_reqDate.SelectedDate is DateTime dt)
            {
                txt_returnRt.Text = _excelService.GenerateReturnRtOrder(dt);
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
                    txt_jobNo.Text = _excelService.GenerateJobNo(reqDate, "RT");
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

        private void Chk_GenReturnRt_Changed(object sender, RoutedEventArgs e)
        {
            if (chk_genReturnRt == null || txt_returnRt == null)
            {
                return;
            }
            if (chk_genReturnRt.IsChecked == true)
            {
                txt_returnRt.IsReadOnly = true;
                UpdateReturnRt();
            }
            else
            {
                txt_returnRt.IsReadOnly = false;
                if (string.IsNullOrWhiteSpace(txt_returnRt.Text))
                {
                    txt_returnRt.Focus();
                }
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
            // 计划表同步必填（仅新增时）
            if (_editTarget == null)
            {
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

            // 新增：同步构造计划记录；编辑：构造关联计划的修改结果（若找到关联计划）
            if (_editTarget == null)
            {
                Plan plan = new()
                {
                    JobNo = txt_jobNo.Text.Trim(),
                    TestItem = cb_testItem.SelectedItem as string,
                    StartDate = dp_startDate.SelectedDate ?? dp_reqDate.SelectedDate,
                    Stage = cb_stage.SelectedItem as string,
                    SampleSize = string.IsNullOrWhiteSpace(txt_sampleSize.Text) ? req.OutQty : txt_sampleSize.Text.Trim(),
                    ModelName = req.ModelName,
                    Product = string.IsNullOrWhiteSpace(txt_product.Text) ? null : txt_product.Text.Trim(),
                    Customer = string.IsNullOrWhiteSpace(txt_customer.Text) ? null : txt_customer.Text.Trim(),
                    Owner = string.IsNullOrWhiteSpace(txt_owner.Text) ? null : txt_owner.Text.Trim(),
                    TestPeriod = string.IsNullOrWhiteSpace(txt_testPeriod.Text) ? null : txt_testPeriod.Text.Trim(),
                    EndDate = dp_endDate.SelectedDate,
                    Status = "Ongoing",
                    CreatedBy = _permission.CurrentUser,
                    CreatedAt = DateTime.Now,
                    UpdatedBy = _permission.CurrentUser,
                    UpdatedAt = DateTime.Now
                };
                if (_db.FreeSql.Select<Plan>().Where(p => p.JobNo == plan.JobNo).Any())
                {
                    _ = MessageBox.Show($"工作編號 [{plan.JobNo}] 已存在", LanguageService.Get("Cap_Info"));
                    return;
                }
                PlanResult = plan;
            }
            else if (_associatedPlan != null)
            {
                // 编辑关联计划（保持 Id/JobNo/创建信息）
                Plan plan = ClonePlan(_associatedPlan);
                plan.TestItem = cb_testItem.SelectedItem as string;
                plan.StartDate = dp_startDate.SelectedDate;
                plan.Stage = cb_stage.SelectedItem as string;
                plan.SampleSize = string.IsNullOrWhiteSpace(txt_sampleSize.Text) ? null : txt_sampleSize.Text.Trim();
                plan.Product = string.IsNullOrWhiteSpace(txt_product.Text) ? null : txt_product.Text.Trim();
                plan.Customer = string.IsNullOrWhiteSpace(txt_customer.Text) ? null : txt_customer.Text.Trim();
                plan.Owner = string.IsNullOrWhiteSpace(txt_owner.Text) ? null : txt_owner.Text.Trim();
                plan.TestPeriod = string.IsNullOrWhiteSpace(txt_testPeriod.Text) ? null : txt_testPeriod.Text.Trim();
                plan.EndDate = dp_endDate.SelectedDate;
                plan.ModelName = req.ModelName;
                plan.UpdatedBy = _permission.CurrentUser;
                plan.UpdatedAt = DateTime.Now;
                PlanResult = plan;
            }

            // 只构造结果，不写数据库；由调用方决定暂存/提审
            DialogResult = true;
        }

        private void Btn_Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
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
