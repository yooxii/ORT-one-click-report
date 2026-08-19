using Microsoft.Extensions.DependencyInjection;
using NLog;
using ORT一键报告.Models;
using ORT一键报告.Services;
using ORT一键报告.Utils;
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;

namespace ORT一键报告.Plans.Views
{
    /// <summary>
    /// WindowPlanEdit.xaml 的交互逻辑：新增/编辑计划记录（计划表基底+领用信息合并为一行）。
    /// 序列号栏支持"自定义输入"与"上传文件"两种方式。
    /// 提审模式（普通用户）：不直接写库，构造结果交由调用方提交审核。
    /// </summary>
    public partial class WindowPlanEdit : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();
        private readonly DatabaseService _db;
        private readonly IPermissionService _permission;
        private readonly Plan _editTarget;
        private readonly bool _submitForReview;

        /// <summary>
        /// 上传方式选择的SN源文件路径
        /// </summary>
        private string _uploadedSnFile;

        /// <summary>
        /// 保存/构造的记录结果（提审模式供调用方提交审核；直存模式为已入库记录）
        /// </summary>
        public Plan PlanResult { get; private set; }

        /// <summary>
        /// 构造新增/编辑对话框
        /// </summary>
        /// <param name="editTarget">编辑目标；null表示新增</param>
        /// <param name="submitForReview">true=提审模式（不写库）</param>
        public WindowPlanEdit(DatabaseService db, IPermissionService permission, Plan editTarget = null, bool submitForReview = false)
        {
            InitializeComponent();
            _db = db;
            _permission = permission;
            _editTarget = editTarget;
            _submitForReview = submitForReview;

            Title = editTarget == null ? "新增计划记录" : "编辑计划记录";
            InitCatalogComboBoxes();
            if (editTarget != null)
            {
                LoadFromPlan(editTarget);
            }
        }

        /// <summary>
        /// 字典选择框初始化：测试项目/产品别/客户别/阶段取自字典表，状况为固定枚举；
        /// 当前值不在字典中时临时加入（避免回显丢失）
        /// </summary>
        private void InitCatalogComboBoxes()
        {
            AdminService admin = App.ServiceProvider.GetRequiredService<AdminService>();
            cb_testItem.ItemsSource = admin.GetTestItems().Select(t => t.Name).ToList();
            cb_product.ItemsSource = admin.GetProducts().Select(p => p.Name).ToList();
            cb_customer.ItemsSource = admin.GetCustomers().Select(c => c.Name).ToList();
            cb_stage.ItemsSource = admin.GetStages().Select(s => s.Name).ToList();
            cb_status.ItemsSource = PlanValidation.ValidStatuses.ToList();
        }

        /// <summary>
        /// 为 ComboBox 设置选中值；值不在字典中时临时补充选项
        /// </summary>
        private static void SetComboValue(System.Windows.Controls.ComboBox combo, string value)
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

        /* ###############################  功能函数  ################################ */

        private void LoadFromPlan(Plan plan)
        {
            // 计划信息
            txt_jobNo.Text = plan.JobNo;
            txt_model.Text = plan.ModelName;
            SetComboValue(cb_testItem, plan.TestItem);
            SetComboValue(cb_stage, plan.Stage);
            SetComboValue(cb_product, plan.Product);
            SetComboValue(cb_customer, plan.Customer);
            txt_sampleSize.Text = plan.SampleSize;
            txt_testPeriod.Text = plan.TestPeriod;
            SetComboValue(cb_status, plan.Status);
            // 领用信息
            txt_reqNo.Text = plan.RequisitionNo;
            txt_returnRt.Text = plan.ReturnRtOrder;
            txt_workOrder.Text = plan.WorkOrder;
            txt_outQty.Text = plan.OutQty;
            txt_lineNo.Text = plan.LineNo;
            txt_reqDate.Text = plan.RequisitionDate;
            // 其他
            txt_owner.Text = plan.Owner;
            txt_remark.Text = plan.Remark;
            // 序列号
            if (!string.IsNullOrWhiteSpace(plan.SnFilePath))
            {
                rb_snFile.IsChecked = true;
                _uploadedSnFile = _db.ResolveAttachmentPath(plan.SnFilePath);
                txt_snFileName.Text = plan.SnFilePath;
            }
            else
            {
                txt_sn.Text = plan.SN;
            }
        }

        /// <summary>
        /// 收集全部字段并校验唯一键（工作編號/领料单据号/回线RT工令，编辑时排除自身），返回false表示校验失败
        /// </summary>
        private bool FillFields(Plan plan)
        {
            long selfId = _editTarget?.Id ?? 0;

            string jobNo = NullIfEmpty(txt_jobNo.Text);
            if (jobNo == null)
            {
                _ = MessageBox.Show("工作編號不能为空", "提示");
                return false;
            }
            // 工作編號格式校验：QRT/RT + 4位年月 + 至少2位编号
            string jobNoError = PlanValidation.ValidateJobNo(jobNo);
            if (jobNoError != null)
            {
                _ = MessageBox.Show(jobNoError, "格式校验失败");
                return false;
            }
            if (_db.FreeSql.Select<Plan>().Where(p => p.JobNo == jobNo && p.Id != selfId).Any())
            {
                _ = MessageBox.Show($"工作編號 [{jobNo}] 已存在，不可重复", "提示");
                return false;
            }
            string reqNo = NullIfEmpty(txt_reqNo.Text);
            if (reqNo != null && _db.FreeSql.Select<Plan>().Where(p => p.RequisitionNo == reqNo && p.Id != selfId).Any())
            {
                _ = MessageBox.Show($"领料单据号 [{reqNo}] 已存在，不可重复", "提示");
                return false;
            }
            string returnRt = NullIfEmpty(txt_returnRt.Text);
            if (returnRt != null && _db.FreeSql.Select<Plan>().Where(p => p.ReturnRtOrder == returnRt && p.Id != selfId).Any())
            {
                _ = MessageBox.Show($"回线RT工令 [{returnRt}] 已存在，不可重复", "提示");
                return false;
            }

            // 计划信息
            plan.JobNo = jobNo;
            plan.ModelName = NullIfEmpty(txt_model.Text);
            plan.TestItem = NullIfEmpty(cb_testItem.SelectedItem as string);
            plan.Stage = NullIfEmpty(cb_stage.SelectedItem as string);
            plan.Product = NullIfEmpty(cb_product.SelectedItem as string);
            plan.Customer = NullIfEmpty(cb_customer.SelectedItem as string);
            plan.SampleSize = NullIfEmpty(txt_sampleSize.Text);
            plan.TestPeriod = NullIfEmpty(txt_testPeriod.Text);
            plan.Status = NullIfEmpty(cb_status.SelectedItem as string);

            // 领用信息
            plan.RequisitionNo = reqNo;
            plan.ReturnRtOrder = returnRt;
            plan.WorkOrder = NullIfEmpty(txt_workOrder.Text);
            plan.OutQty = NullIfEmpty(txt_outQty.Text);
            plan.LineNo = NullIfEmpty(txt_lineNo.Text);
            plan.RequisitionDate = NullIfEmpty(txt_reqDate.Text);

            plan.Owner = NullIfEmpty(txt_owner.Text);
            plan.Remark = NullIfEmpty(txt_remark.Text);
            return true;
        }

        /// <summary>
        /// 将上传的SN文件复制到附件目录，命名 {简短日期}_{单据号/工作编号}_{机种名称}_{原文件名}，返回保存的文件名
        /// </summary>
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

        private static string NullIfEmpty(string s) => string.IsNullOrWhiteSpace(s) ? null : s.Trim();

        private static string Clean(string name)
        {
            string cleaned = Regex.Replace(name ?? "", $"[{Regex.Escape(new string(Path.GetInvalidFileNameChars()))}]", "_").Trim();
            return cleaned == "" ? "_" : cleaned;
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

        private void Btn_Save_Click(object sender, RoutedEventArgs e)
        {
            // 编辑时以原记录为基底（保留Id/审计/日期解析等未编辑字段），新增时创建新记录
            Plan plan = _editTarget == null
                ? new Plan { CreatedBy = _permission.CurrentUser, CreatedAt = DateTime.Now }
                : ClonePlan(_editTarget);
            plan.UpdatedBy = _permission.CurrentUser;
            plan.UpdatedAt = DateTime.Now;

            if (!FillFields(plan))
            {
                return;
            }

            // 序列号：自定义输入 或 上传文件（复制到附件目录）
            if (rb_snInput.IsChecked == true)
            {
                plan.SN = NullIfEmpty(txt_sn.Text);
                if (_editTarget == null || rb_snInput.IsChecked == true)
                {
                    // 改为文本输入时清除原附件引用
                    if (_editTarget != null)
                    {
                        plan.SnFilePath = null;
                    }
                }
            }
            else
            {
                if (_uploadedSnFile == null)
                {
                    _ = MessageBox.Show("请先选择要上传的序列号文件", "提示");
                    return;
                }
                // 编辑且附件未变化时保留原文件名
                string existingFile = _editTarget == null ? null : _db.ResolveAttachmentPath(_editTarget.SnFilePath);
                if (_editTarget != null && string.Equals(_uploadedSnFile, existingFile, StringComparison.OrdinalIgnoreCase))
                {
                    plan.SnFilePath = _editTarget.SnFilePath;
                }
                else
                {
                    if (!File.Exists(_uploadedSnFile))
                    {
                        _ = MessageBox.Show($"序列号文件不存在:\n{_uploadedSnFile}", "提示");
                        return;
                    }
                    string key = plan.JobNo ?? plan.RequisitionNo;
                    string savedName = SaveSnFile(_uploadedSnFile, key, plan.ModelName);
                    if (savedName == null)
                    {
                        return;
                    }
                    plan.SnFilePath = savedName;
                }
                plan.SN = plan.SN; // 保留原SN文本
            }

            if (_submitForReview)
            {
                // 提审模式：不写库，返回构造结果
                PlanResult = plan;
                DialogResult = true;
                return;
            }

            try
            {
                if (_editTarget == null)
                {
                    _db.FreeSql.Insert(plan).ExecuteAffrows();
                    _logger.Info($"新增计划记录: 机种={plan.ModelName}, 工作編號={plan.JobNo}");
                }
                else
                {
                    _db.FreeSql.Update<Plan>().SetSource(plan).Where(p => p.Id == plan.Id).ExecuteAffrows();
                    _logger.Info($"编辑计划记录: Id={plan.Id}, 机种={plan.ModelName}");
                }
                PlanResult = plan;
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

        /// <summary>
        /// 深拷贝计划记录（编辑基底，避免直接修改列表中的对象）
        /// </summary>
        private static Plan ClonePlan(Plan source)
            => Newtonsoft.Json.JsonConvert.DeserializeObject<Plan>(
                Newtonsoft.Json.JsonConvert.SerializeObject(source));
    }
}
