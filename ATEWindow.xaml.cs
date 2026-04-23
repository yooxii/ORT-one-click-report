using Microsoft.Win32;
using OfficeOpenXml;
using System;
using NLog;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.IO;
using System.Data;
using static ORT一键报告.ReportHeader;

namespace ORT一键报告
{
    /// <summary>
    /// ATEWindow.xaml 的交互逻辑
    /// </summary>
    public partial class ATEWindow : Window
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();



        public class ATEItem
        {
            public string ItemName { get; set; }
            public string OutputType { get; set; }
            public string SN { get; set; }
            public string Value { get; set; }
            public string MaxSpec { get; set; }
            public string MinSpec { get; set; }
            public bool IsPassed
            {
                get
                {
                    try
                    {
                        if (MaxSpec[0] == '-')
                        {
                            double _Max = double.Parse(MaxSpec);
                            double _Min = double.Parse(MinSpec);
                            MaxSpec = Math.Max(_Max, _Min).ToString();
                            MinSpec = Math.Min(_Max, _Min).ToString();
                        }
                        if ((MaxSpec == "*" || MaxSpec == " " || double.Parse(Value) <= double.Parse(MaxSpec)) &&
                            (MinSpec == "*" || MinSpec == " " || double.Parse(Value) >= double.Parse(MinSpec)))
                        {
                            return true;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogManager.GetCurrentClassLogger().Error(ex, "数据转换错误");
                    }
                    return false;
                }
            }
        }
        public class ATERows
        {
            public string SN { get; set; }
            public List<string> Datas { get; set; }
            public List<int> BadDataIndexs { get; set; }
            public void Clear()
            {
                SN = "";
                Datas?.Clear();
                BadDataIndexs?.Clear();
            }
        }
        public class ATEItems
        {
            public List<string> ItemNames { get; set; }
            public List<string> OutputTypes { get; set; }
            public List<string> ShortTitles { get; set; }
            public List<ATERows> BeforeDatas { get; set; }
            public List<ATERows> AfterDatas { get; set; }
            public List<string> MaxSpecs { get; set; }
            public List<string> MinSpecs { get; set; }
            public int Count;
            public ATEItems()
            {
                BeforeDatas = new List<ATERows>();
                AfterDatas = new List<ATERows>();
                ShortTitles = new List<string>();
                Count = 0;
            }
            public void Clear()
            {
                ItemNames?.Clear();
                OutputTypes?.Clear();
                ShortTitles?.Clear();
                BeforeDatas?.Clear();
                AfterDatas?.Clear();
                MaxSpecs?.Clear();
                MinSpecs?.Clear();
                Count = 0;
            }
            public void Add(ATEItem item, bool isBefore)
            {
                List<ATERows> Datas = isBefore ? BeforeDatas : AfterDatas;
                if (Count == 0)
                {
                    // 首次新建
                    Datas.Add(new ATERows()
                    {
                        SN = item.SN,
                        Datas = new List<string>() { item.Value },
                        BadDataIndexs = item.IsPassed ? new List<int>() : new List<int>() { 0 }
                    });
                    ItemNames = new List<string>() { item.ItemName };
                    OutputTypes = new List<string>() { item.OutputType };
                    MaxSpecs = new List<string>() { item.MaxSpec };
                    MinSpecs = new List<string>() { item.MinSpec };
                }
                else
                {
                    // 每个SN对应数据源的一行
                    ATERows tempDatas = Datas.Find(t => t.SN == item.SN);
                    if (tempDatas != null)
                    {
                        if (!item.IsPassed)
                        {
                            tempDatas.BadDataIndexs.Add(Datas.Count);
                        }
                        tempDatas.Datas.Add(item.Value);
                    }
                    else
                    {
                        Datas.Add(new ATERows()
                        {
                            SN = item.SN,
                            Datas = new List<string>() { item.Value },
                            BadDataIndexs = item.IsPassed ? new List<int>() : new List<int>() { 0 }
                        });
                    }
                    // 测试名称和输出类型如果不同就新增
                    if (ItemNames.FindIndex(t => t.Equals(item.ItemName)) == -1 || OutputTypes.FindIndex(t => t.Equals(item.OutputType)) == -1)
                    {
                        ItemNames.Add(item.ItemName);
                        OutputTypes.Add(item.OutputType);
                        MaxSpecs.Add(item.MaxSpec);
                        MinSpecs.Add(item.MinSpec);
                    }
                }
                Count++;
            }
            public DataTable ToItemSource()
            {
                if (ItemNames is null)
                {
                    return new DataTable();
                }
                DataTable res = new DataTable();
                res.Columns.Add("SN", typeof(string));
                for (int i = 0; i < ItemNames.Count; i++)
                {
                    string title = $"{ItemNames[i].Split(' ')[0]} {OutputTypes[i]}";
                    ShortTitles.Add(title);
                    res.Columns.Add(title, typeof(string));
                }
                foreach (ATERows bData in BeforeDatas)
                {
                    List<string> tmp = new List<string>() { bData.SN };
                    tmp.AddRange(bData.Datas);
                    res.Rows.Add(tmp.ToArray());
                }
                foreach (ATERows aData in AfterDatas)
                {
                    List<string> tmp = new List<string>() { aData.SN };
                    tmp.AddRange(aData.Datas);
                    res.Rows.Add(tmp.ToArray());
                }
                List<string> tmpSpec = new List<string>() { "MAX_SPEC" };
                tmpSpec.AddRange(MaxSpecs);
                res.Rows.Add(tmpSpec.ToArray());

                tmpSpec.Clear();
                tmpSpec.Add("MIN_SPEC");
                tmpSpec.AddRange(MinSpecs);
                res.Rows.Add(tmpSpec.ToArray());

                return res;
            }
        }


        public ATEItems ATEitems = new ATEItems();
        public DataTable ATETable;
        public string ATEFilePath;
        public string SaveATEPath;
        public string TempATEDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ORTTemp", "ATE");

        public ATEWindow()
        {
            InitializeComponent();
            Closed += ATEWindow_Closed;
        }
        private void ATEWindow_Closed(object sender, EventArgs e)
        {
            ClearTempDir(_logger);
        }

        /* ###############################  功能函数  ################################ */
        private string F_xls2xlsx(string filePath)
        {
            // 1. 基础校验
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return null;
            }

            // 2. 规范扩展名判断
            string extension = System.IO.Path.GetExtension(filePath);
            if (!extension.Equals(".xls", StringComparison.OrdinalIgnoreCase))
            {
                return filePath; // 已经是 xlsx 或其他格式，直接返回
            }

            if (!File.Exists(filePath))
            {
                _logger.Error($"文件不存在: {filePath}");
                return null;
            }

            string outputFilePath = null;
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Microsoft.Office.Interop.Excel.Workbook workbook = null;

            try
            {
                // 3. 生成唯一的临时输出路径 (防止文件名冲突)
                if (!Directory.Exists(TempATEDir))
                {
                    _ = Directory.CreateDirectory(TempATEDir);
                }
                outputFilePath = System.IO.Path.Combine(TempATEDir, Guid.NewGuid().ToString() + ".xlsx");

                // 4. 启动 Excel (增加错误处理)
                excelApp = new Microsoft.Office.Interop.Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                // 5. 打开文件 (增加重试机制以应对文件被占用)
                int retryCount = 0;
                bool opened = false;
                while (retryCount < 3 && !opened)
                {
                    try
                    {
                        // 使用 ReadOnly 模式打开，减少文件锁冲突
                        workbook = excelApp.Workbooks.Open(filePath, ReadOnly: true);
                        opened = true;
                    }
                    catch (IOException)
                    {
                        retryCount++;
                        System.Threading.Thread.Sleep(500); // 等待 0.5 秒后重试
                    }
                }

                if (!opened) throw new Exception("无法打开源文件，可能文件正被占用。");

                // 6. 另存为 xlsx (FileFormat 51 = xlOpenXMLWorkbook)
                workbook.SaveAs(outputFilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook);
                workbook.Close(false); // 关闭源文件，不保存更改

                return outputFilePath;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "xls文件转xlsx文件失败");
                // 如果转换失败，确保删除可能产生的半成品文件
                if (outputFilePath != null && File.Exists(outputFilePath))
                {
                    try { File.Delete(outputFilePath); } catch { }
                }
                return null;
            }
            finally
            {
                // 7. 严格的 COM 对象释放
                if (workbook != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                }

                if (excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }

                // 强制垃圾回收，帮助释放 COM 引用
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private bool IsPair(string a, string b, char aC = '1', char bC = '2')
        {
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b) || a.Length != b.Length)
            {
                return false;
            }
            int diffCount = 0;
            int diffIndex = -1;
            for (int i = 0; i < a.Length; i++)
            {
                if (a[i] != b[i])
                {
                    diffCount++;
                    diffIndex = i;
                    if (diffCount > 1)
                    {
                        return false;
                    }
                }
            }
            if (diffCount != 1)
            {
                return false;
            }
            char c1 = a[diffIndex];
            char c2 = b[diffIndex];
            return (c1 == bC && c2 == aC) || (c1 == aC && c2 == bC);
        }

        private List<string> FindSNs(ExcelWorksheet ws, DataCell startCell, DataCell maxSpecCell)
        {
            List<string> SNs = new List<string>();
            for (int r = startCell.Row + 1; r < maxSpecCell.Row; r++)
            {
                SNs.Add(ws.Cells[r, 1].Text);
            }
            return SNs;
        }

        private void ReadATEDatas(string fileName)
        {
            try
            {
                FileInfo ateFile = new FileInfo(System.IO.Path.GetFullPath(fileName));
                if (!ateFile.Exists)
                {
                    throw new FileNotFoundException($"ATE数据文件在{fileName}未找到");
                }
                ExcelPackage atePackage = new ExcelPackage(ateFile);
                ExcelWorkbook wb = atePackage.Workbook;
                ExcelWorksheet ws = wb.Worksheets[0];

                DataCell startCell = FindCellByValue(ws, "s/n");
                DataCell maxSpecCell = FindCellByValue(ws, "MAX_SPEC");

                if (startCell.Row >= maxSpecCell.Row)
                {
                    _logger.Warn("无ATE数据");
                    _ = MessageBox.Show("无ATE数据");
                    return;
                }

                List<string> SNs = FindSNs(ws, startCell, maxSpecCell);
                int flag = IsPair(SNs[0], SNs[1]) ? 1 : IsPair(SNs[0], SNs[SNs.Count / 2]) ? 2 : 3;
                if (ATEitems.Count != 0)
                {
                    ATEitems.Clear();
                }
                for (int r = startCell.Row + 1; r < maxSpecCell.Row; r++)
                {
                    for (int c = startCell.Column + 1; c <= ws.Dimension.End.Column; c++)
                    {
                        if (ws.Cells[startCell.Row, c].Text == "")
                        {
                            break;
                        }
                        bool isBefore = flag == 1 ? (r - startCell.Row) % 2 == 1 : flag != 2 || (r - startCell.Row) / (SNs.Count / 2) == 0;
                        ATEitems.Add(new ATEItem()
                        {
                            SN = ws.Cells[r, startCell.Column].Text,
                            Value = ws.Cells[r, c].Text,
                            MaxSpec = ws.Cells[maxSpecCell.Row, c].Text,
                            MinSpec = ws.Cells[maxSpecCell.Row + 1, c].Text,
                            ItemName = ws.Cells[startCell.Row - 1, c].Text,
                            OutputType = ws.Cells[startCell.Row, c].Text
                        }, isBefore);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "读取ATE数据发生错误");
                _ = MessageBox.Show(ex + "读取ATE数据发生错误", "错误");
            }
        }

        private void MakeATEDatas(FileInfo ATEtemp)
        {
            if (string.IsNullOrEmpty(ATEFilePath))
            {
                return;
            }
            _logger.Info("ATE报告生成中...");
            try
            {
                ExcelPackage package = new ExcelPackage(ATEtemp);
                ExcelWorkbook wb = package.Workbook;
                ExcelWorksheet ws = wb.Worksheets[0];

                // 1. 写入标题和Spec
                var sourceRange = ws.Cells[1, 4, ws.Dimension.End.Row, 4];
                for (int c = 0; c < ATEitems.ItemNames.Count; c++)
                {
                    ws.Cells[3, c + 4].Value = ATEitems.ItemNames[c];
                    ws.Cells[4, c + 4].Value = ATEitems.OutputTypes[c];
                    ws.Cells[11, c + 4].Value = ATEitems.MaxSpecs[c];
                    ws.Cells[12, c + 4].Value = ATEitems.MinSpecs[c];
                    if (c > 0)
                    {
                        // 2. 复制样式和公式
                        sourceRange.CopyStyles(ws.Cells[1, c + 4, ws.Dimension.End.Row, c + 4]);
                        sourceRange.CopyFormulas(ws.Cells[1, c + 4, ws.Dimension.End.Row, c + 4]);
                    }
                }
                // 3. 编辑行数量
                Make_EditRowCounts(ws);
                // 4. 写入数据
                for (int r = 0; r < ATETable.Rows.Count; r++)
                {
                    List<object> dataArray = ATETable.Rows[r].ItemArray.ToList();
                    for (int c = 0; c < dataArray.Count; c++)
                    {
                        try
                        {
                            ws.Cells[r + 5, c + 3].Value = double.Parse(dataArray[c].ToString());
                        }
                        catch
                        {
                            ws.Cells[r + 5, c + 3].Value = dataArray[c].ToString();
                        }
                    }
                }
                Make_SaveAsExcel(package);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "ATE报告生成失败");
                _ = MessageBox.Show(ex.Message, "ATE报告生成失败");
            }
        }

        private void Make_EditRowCounts(ExcelWorksheet ws)
        {
            int addRowCount = 0;
            // 试验前
            if (ATEitems.BeforeDatas.Count > 3)
            {
                addRowCount += ATEitems.BeforeDatas.Count - 3;
                ws.InsertRow(6, addRowCount);
                ws.Cells[5, 1, 5, ws.Dimension.End.Column].CopyStyles(ws.Cells[6, 1, 6 + addRowCount, ws.Dimension.End.Column]);
                ws.Rows[6, 6 + addRowCount].Height = ws.Row(5).Height;
            }
            else if (ATEitems.BeforeDatas.Count < 3)
            {
                ws.DeleteRow(5 + addRowCount, 3 - ATEitems.BeforeDatas.Count);
                addRowCount -= 3 - ATEitems.BeforeDatas.Count;
            }
            // 试验后
            if (ATEitems.AfterDatas.Count > 3)
            {
                ws.InsertRow(9 + addRowCount, ATEitems.AfterDatas.Count - 3);
                ws.Cells[8 + addRowCount, 1, 8 + addRowCount, ws.Dimension.End.Column].CopyStyles(ws.Cells[8 + addRowCount, 1, 8 + addRowCount + ATEitems.AfterDatas.Count - 3, ws.Dimension.End.Column]);
                ws.Rows[8 + addRowCount, 8 + addRowCount + ATEitems.AfterDatas.Count - 3].Height = ws.Row(5).Height;
                addRowCount += ATEitems.AfterDatas.Count - 3;
            }
            else if (ATEitems.AfterDatas.Count < 3)
            {
                ws.DeleteRow(8 + addRowCount, 3 - ATEitems.AfterDatas.Count);
                addRowCount -= 3 - ATEitems.AfterDatas.Count;
                // 如果没有试验后数据，删除试验后的部分
                if (ATEitems.AfterDatas.Count == 0)
                {
                    ws.DeleteRow(17 + addRowCount, 4);
                }
            }
        }

        private void Make_SaveAsExcel(ExcelPackage package)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                FileName = System.IO.Path.GetFileName(ATEFilePath),
                Filter = "Excel文件|*.xlsx;*.xls",
                InitialDirectory = System.IO.Path.GetDirectoryName(ATEFilePath)
            };
            SaveATEPath = saveFileDialog.ShowDialog() == true
                ? saveFileDialog.FileName
                : System.IO.Path.Combine(Directory.GetCurrentDirectory(), System.IO.Path.GetFileName(ATEFilePath));
            package.SaveAs(SaveATEPath);
            _logger.Info($"ATE报告已保存，路径：{SaveATEPath}");
            _ = MessageBox.Show($"ATE报告已保存，路径：{SaveATEPath}");
        }

        private FileInfo GetATETemplate()
        {
            FileInfo ATEtemp;
            try
            {
                ATEtemp = new FileInfo(System.IO.Path.GetFullPath(text_ATETemplate.Text));
                if (!ATEtemp.Exists)
                {
                    throw new FileNotFoundException("ATE模板路径不存在");
                }
            }
            catch
            {
                string currentPath = Directory.GetCurrentDirectory();
                string ATEtemplatePath = GetTemplatePath(System.IO.Path.Combine(currentPath, "Templates"), "ATE");
                if (File.Exists(ATEtemplatePath))
                {
                    text_ATETemplate.Text = ATEtemplatePath;
                    ATEtemp = new FileInfo(ATEtemplatePath);
                }
                else
                {
                    throw new FileNotFoundException("默认ATE模板不存在，请主动选择ATE模板路径");
                }
            }
            return ATEtemp;
        }

        /* ###############################  事件函数  ################################ */
        private async void OpenATEDatas_Click(object sender, RoutedEventArgs e)
        {
            FileDialog ateDialog = new OpenFileDialog()
            {
                Filter = "ATE数据文件|*.xls;*.xlsx",
            };
            if (ateDialog.ShowDialog() == true)
            {
                PopupWindow popup = new PopupWindow() { Title = "处理中", Message = "请耐心等待..." };

                try
                {
                    popup.Show();
                    await Task.Run(() =>
                    {
                        ATEFilePath = F_xls2xlsx(ateDialog.FileName);

                        if (string.IsNullOrEmpty(ATEFilePath))
                        {
                            _logger.Warn("ATE数据路径选择错误");
                            _ = MessageBox.Show("ATE数据路径选择错误", "错误");
                            throw new Exception("ATE数据路径选择错误");
                        }

                        ReadATEDatas(ATEFilePath);
                        ATETable = ATEitems.ToItemSource();
                    });
                    dataGridATE.ItemsSource = ATETable.DefaultView;
                }
                finally
                {
                    popup.Close();
                }
            }
            else
            {
                _logger.Warn("未选择ATE数据文件");
                _ = MessageBox.Show("未选择ATE数据文件");
            }
        }

        private async void SaveATEDatas_Click(object sender, RoutedEventArgs e)
        {
            FileInfo ATEtemp = GetATETemplate();
            await Task.Run(() => { MakeATEDatas(ATEtemp); });
        }

        private void btn_ATETemplate_Click(object sender, RoutedEventArgs e)
        {
            FileDialog ATEtemplate = new OpenFileDialog()
            {
                Filter = "ATE数据模板|*.xls;*.xlsx",
            };
            if (ATEtemplate.ShowDialog() == true)
            {
                text_ATETemplate.Text = ATEtemplate.FileName;
            }
            else
            {
                string currentPath = Directory.GetCurrentDirectory();
                string ATEtemplatePath = GetTemplatePath(System.IO.Path.Combine(currentPath, "Templates"), "ATE");
                if (File.Exists(ATEtemplatePath))
                {
                    text_ATETemplate.Text = ATEtemplatePath;
                }
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void btn_reread_Click(object sender, RoutedEventArgs e)
        {
            ReadATEDatas(ATEFilePath);
            ATETable = ATEitems.ToItemSource();
            dataGridATE.ItemsSource = ATETable.DefaultView;
        }

        private void MakeATEReport_Click(object sender, RoutedEventArgs e)
        {
            Thickness btn_thick = new Thickness(10, 1, 10, 1);
            Button btn_before = new Button() { Content = "试验前", Margin = btn_thick, Width = 60, Height = 25 };
            Button btn_after = new Button() { Content = "试验后", Margin = btn_thick, Width = 60, Height = 25 };
            Button btn_both = new Button() { Content = "都要", Margin = btn_thick, Width = 60, Height = 25 };
            //PopupWindow popup = new PopupWindow("保存选项", "ATE要保存哪部分数据？", new List<object> { btn_before, btn_after, btn_both });

        }
    }
}
