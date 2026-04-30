using Microsoft.Win32;
using NLog;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace ORT一键报告
{
    //2.1 Conducted EMI Measurement
    class EMIUUTSetup
    {
        public List<string> SN { get; set; } = new List<string>();
        public List<string> Voltage { get; set; } = new List<string>();
        public List<string> Load { get; set; } = new List<string>();
        public List<string> LISN { get; set; } = new List<string>();
    }

    /// <summary>
    /// EMIReportPage.xaml 的交互逻辑
    /// </summary>
    public partial class EMIReportPage : UserControl
    {
        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        private EMIUUTSetup emiUUTsetup = new EMIUUTSetup();
        private List<string> emiDocxFiles = new List<string>();
        private List<string> emiPdfFiles = new List<string>();
        public EMIReportPage()
        {
            InitializeComponent();
        }

        /* ###############################  功能函数  ################################ */
        private int[] ConfirmOrder(string datafile)
        {
            try
            {
                string[] infos = Path.GetFileNameWithoutExtension(datafile).Split('-');
                int[] res = new int[4];
                for (int i = 0; i < infos.Length; i++)
                {
                    if (infos[i].Length <= 1)
                    {
                        res[3] = i;
                    }
                    else if (infos[i].Length <= 4)
                    {
                        if (infos[i].Contains("V"))
                        {
                            res[1] = i;
                        }
                        else if (infos[i].Contains("%"))
                        {
                            res[2] = i;
                        }
                        else
                        {
                            res[0] = i;
                        }
                    }
                    else
                    {
                        res[0] = i;
                    }
                }
                return res;
            }
            catch
            {
                return new int[4] { 0, 1, 2, 3 };
            }
        }

        private void ReadPath(string dataDir)
        {
            int[] order = ConfirmOrder(dataDir);
            emiDocxFiles.Clear();
            foreach (string datafile in Directory.GetFiles(dataDir))
            {
                if (Path.GetExtension(datafile).ToLower().Contains("docx"))
                {
                    string[] infos = Path.GetFileNameWithoutExtension(datafile).Split('-');
                    if (!emiUUTsetup.SN.Contains(infos[order[0]])) emiUUTsetup.SN.Add(infos[order[0]]);
                    if (!emiUUTsetup.Voltage.Contains(infos[order[1]])) emiUUTsetup.Voltage.Add(infos[order[1]]);
                    if (!emiUUTsetup.Load.Contains(infos[order[2]])) emiUUTsetup.Load.Add(infos[order[2]]);
                    if (!emiUUTsetup.LISN.Contains(infos[order[3]])) emiUUTsetup.LISN.Add(infos[order[3]]);
                    emiDocxFiles.Add(datafile);
                }
                else if (Path.GetExtension(datafile).ToLower().Contains("pdf"))
                {
                    emiPdfFiles.Add(datafile);
                }
            }
        }

        public void ConvertWord2Pdf(string WordPath)
        {
            if (WordPath is null || WordPath == "")
            {
                return;
            }
            if (Path.GetExtension(WordPath) == "docx")
            {
                Docx2Pdf.ConvertToPdf(WordPath);
            }
            else if (Directory.Exists(WordPath))
            {
                ReadPath(WordPath);
                if (emiDocxFiles == null || emiDocxFiles.Count == 0)
                {
                    _logger.Error("给定的目录路径错误或不存在docx文件");
                    return;
                }
                foreach (string file in emiDocxFiles)
                {
                    Docx2Pdf.ConvertToPdf(file);
                }
            }
            else
            {
                _logger.Error("给定的docx文件路径错误或不存在");
            }
        }

        /* ###############################  事件函数  ################################ */

        private void btn_Dir_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string textbox)
            {
                if (FindName(textbox) is TextBox textBox)
                {
                    FileDialog fd = new OpenFileDialog()
                    {
                        Filter = "EMI数据文件|*.pdf;*.docx|所有文件|*.*"
                    };
                    _ = fd.ShowDialog();
                    if (fd.FileName is string fileName && fileName != "")
                    {
                        textBox.Text = Path.GetDirectoryName(fd.FileName);
                    }
                }
            }
        }

        private void btn_File_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string textbox)
            {
                if (FindName(textbox) is TextBox textBox)
                {
                    FileDialog fd = new OpenFileDialog()
                    {
                        Filter = "EMI模版文件|*.xlsx;*.xlsx|所有文件|*.*"
                    };
                    _ = fd.ShowDialog();
                    if (fd.FileName is string fileName && fileName != "")
                    {
                        textBox.Text = fd.FileName;
                    }
                }
            }
        }

        private void btnReadPath_Click(object sender, RoutedEventArgs e)
        {
            string dataDir = DataPath_text.Text;
            if (!Directory.Exists(dataDir))
            {
                return;
            }
            ReadPath(dataDir);
        }

        private void btnToPDF_Click(object sender, RoutedEventArgs e)
        {
            ConvertWord2Pdf(DataPath_text.Text);
        }
    }
}
