using Microsoft.Extensions.DependencyInjection;
using NLog;
using ORT一键报告.Models;
using ORT一键报告.Services;
using System;
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
    /// </summary>
    public partial class WindowRequisitionEdit : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;
        private readonly IPermissionService _permission;
        private readonly AdminService _admin;
        private readonly PlanExcelService _excelService;
        private readonly Requisition _editTarget;
        private readonly bool _submitForReview;

        /// <summary>
        /// 上传方式选择的SN源文件路径
        /// </summary>
        private string _uploadedSnFile;

        /// <summary>
        /// 保存/构造的领退记录结果（提审模式供调用方提交审核）
        /// </summary>
        public Requisition RequisitionResult { get; private set; }

        /// <summary>
        /// 保存/构造的计划记录结果（同步新增计划）
        /// </summary>
        public Plan PlanResult { get; private set; }

        public WindowRequisitionEdit(DatabaseService db, IPermissionService permission, AdminService admin,
            PlanExcelService excelService, Requisition editTarget = null, bool submitForReview = false)
        {
            InitializeComponent();
            _db = db;
            _permission = permission;
            _admin = admin;
            _excelService = excelService;
            _editTarget = editTarget;
            _submitForReview = submitForReview;

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
            }
        }

        /* ###############################  加载  ################################ */

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
        /// 从 Work Order 自动补全 D/C（倒数第三位起的两位）与 線別（倒数第六位起的三位）
        /// </summary>
        private void AutoFillFromWorkOrder()
        {
            string wo = txt_workOrder?.Text?.Trim();
            if (string.IsNullOrWhiteSpace(wo) || wo.Length < 6)
            {
                return;
            }
            // 線別：倒数第六位起的三位字符串
            txt_lineNo.Text = wo.Substring(wo.Length - 6, 3);
            // D/C：倒数第三位起的两位
            if (wo.Length >= 3)
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
        /// 刷新计划表同步信息（工作编号/样品数/产品别/客户别/负责人/试验时间/结束日期）
        /// </summary>
        private void UpdateAutoPlan()
        {
            if (_editTarget != null)
            {
                return; // 编辑模式不同步新增计划
            }
            if (dp_reqDate.SelectedDate is not DateTime reqDate)
            {
                return;
            }
            // 开始时间默认与领用日期一致
            if (dp_startDate.SelectedDate == null || dp_startDate.SelectedDate != dp_reqDate.SelectedDate)
            {
                if (dp_startDate.SelectedDate == null)
                {
                    dp_startDate.SelectedDate = dp_reqDate.SelectedDate;
                }
            }
            // 工作编号：RT{当前年月}{编号}
            txt_jobNo.Text = _excelService.GenerateJobNo(reqDate, "RT");
            // 样品数 = 领出数量
            txt_sampleSize.Text = txt_outQty?.Text?.Trim();

            // 产品别/客户别：根据机种名查询
            string model = txt_model?.Text?.Trim();
            if (!string.IsNullOrWhiteSpace(model))
            {
                ModelMapping mapping = _admin.FindModelMapping(model);
                txt_product.Text = mapping?.Product ?? "";
                txt_customer.Text = mapping?.Customer ?? "";
            }
            // 负责人/试验时间/结束日期：根据测试项目查询
            string testItem = cb_testItem.SelectedItem as string;
            if (!string.IsNullOrWhiteSpace(testItem))
            {
                TestItemCatalog item = _admin.GetTestItems().FirstOrDefault(t => t.Name == testItem);
                if (item != null)
                {
                    txt_owner.Text = item.Owner;
                    txt_testPeriod.Text = item.Period;
                    if (int.TryParse(item.Period, out int hours))
                    {
                        DateTime start = dp_startDate.SelectedDate ?? reqDate;
                        dp_endDate.SelectedDate = start.AddHours(hours);
                    }
                }
            }
        }

        /* ###############################  事件函数  ################################ */

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
                Title = "选择序列号文件",
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
                txt_returnRt.Text = "";
                txt_returnRt.Focus();
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
                _ = MessageBox.Show("請填寫領用日期", "提示");
                return;
            }
            if (string.IsNullOrWhiteSpace(txt_reqNo.Text))
            {
                _ = MessageBox.Show("請填寫領料單据號", "提示");
                return;
            }
            if (string.IsNullOrWhiteSpace(txt_model.Text))
            {
                _ = MessageBox.Show("請填寫機種名稱", "提示");
                return;
            }
            if (string.IsNullOrWhiteSpace(txt_outQty.Text))
            {
                _ = MessageBox.Show("請填寫領出數量", "提示");
                return;
            }
            if (string.IsNullOrWhiteSpace(txt_rev.Text))
            {
                _ = MessageBox.Show("請填寫REV.", "提示");
                return;
            }
            if (string.IsNullOrWhiteSpace(txt_workOrder.Text))
            {
                _ = MessageBox.Show("請填寫Work Order", "提示");
                return;
            }
            if (rb_snInput.IsChecked == true && string.IsNullOrWhiteSpace(txt_sn.Text))
            {
                _ = MessageBox.Show("請填寫S/N或上傳附件", "提示");
                return;
            }
            if (rb_snFile.IsChecked == true && _uploadedSnFile == null)
            {
                _ = MessageBox.Show("請先選擇要上傳的序列號文件", "提示");
                return;
            }
            // 计划表同步必填（新增时）
            if (_editTarget == null)
            {
                if (cb_testItem.SelectedItem == null)
                {
                    _ = MessageBox.Show("請選擇測試項目", "提示");
                    return;
                }
                if (cb_stage.SelectedItem == null)
                {
                    _ = MessageBox.Show("請選擇階段", "提示");
                    return;
                }
            }

            // 唯一键校验
            long selfId = _editTarget?.Id ?? 0;
            if (_db.FreeSql.Select<Requisition>().Where(r => r.RequisitionNo == txt_reqNo.Text.Trim() && r.Id != selfId).Any())
            {
                _ = MessageBox.Show($"領料單据號 [{txt_reqNo.Text.Trim()}] 已存在", "提示");
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

            // 同步计划（仅新增时）
            if (_editTarget == null)
            {
                Plan plan = new()
                {
                    JobNo = txt_jobNo.Text.Trim(),
                    TestItem = cb_testItem.SelectedItem as string,
                    StartDate = dp_startDate.SelectedDate ?? dp_reqDate.SelectedDate,
                    Stage = cb_stage.SelectedItem as string,
                    SampleSize = req.OutQty,
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
                    _ = MessageBox.Show($"工作編號 [{plan.JobNo}] 已存在", "提示");
                    return;
                }
                PlanResult = plan;
            }

            if (_submitForReview)
            {
                DialogResult = true;
                return;
            }

            try
            {
                if (_editTarget == null)
                {
                    req.Id = _db.FreeSql.Insert(req).ExecuteIdentity();
                    _db.FreeSql.Insert(PlanResult).ExecuteAffrows();
                    _logger.Info($"新增领退[{req.RequisitionNo}]并同步计划[{PlanResult.JobNo}]");
                }
                else
                {
                    _db.FreeSql.Update<Requisition>().SetSource(req).Where(r => r.Id == req.Id).ExecuteAffrows();
                    _logger.Info($"编辑领退: Id={req.Id}, {req.RequisitionNo}");
                }
                DialogResult = true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "保存领退记录失败");
                _ = MessageBox.Show($"保存失败:\n{ex.Message}", "错误");
            }
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
                _ = MessageBox.Show($"保存上传文件失败:\n{ex.Message}", "错误");
                return null;
            }
        }

        private static Requisition CloneReq(Requisition source)
            => Newtonsoft.Json.JsonConvert.DeserializeObject<Requisition>(
                Newtonsoft.Json.JsonConvert.SerializeObject(source));

        private static string Clean(string name)
        {
            string cleaned = Regex.Replace(name ?? "", $"[{Regex.Escape(new string(Path.GetInvalidFileNameChars()))}]", "_").Trim();
            return cleaned == "" ? "_" : cleaned;
        }
    }
}
