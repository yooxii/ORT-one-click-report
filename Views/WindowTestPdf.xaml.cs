using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using Microsoft.Win32;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;

namespace ORT一键报告
{
    /// <summary>
    /// TestPdfWindow.xaml 的交互逻辑
    /// </summary>
    public partial class TestPdfWindow : Window
    {
        public TestPdfWindow()
        {
            InitializeComponent();
        }

        /* ###############################  功能函数  ################################ */

        public string DealPdfFile(string pdfPath)
        {
            PdfReader reader = new PdfReader(pdfPath);
            PdfDocument pdfDoc = new PdfDocument(reader);

            PdfPage page1 = pdfDoc.GetPage(1);
            PdfPage page2 = pdfDoc.GetPage(2);
            string textRenderInfos = $"{PdfTextExtractor.GetTextFromPage(page1)}\n{PdfTextExtractor.GetTextFromPage(page2)}";

            return textRenderInfos;
        }

        /* ###############################  事件函数  ################################ */

        private void Btn_PDFPath_Click(object sender, RoutedEventArgs e)
        {
            FileDialog fd = new OpenFileDialog()
            {
                Filter = "PDF文件|*.pdf|docx文件|*.docx"
            };
            _ = fd.ShowDialog();
            if (fd.FileName is string fileName && fileName != "")
            {
                TextBox_PdfPath.Text = fileName;
                string rescsv = SmartTableExtractor.ConvertWordTablesToCsv(fileName);
                TextBlock_PdfContent.Text = rescsv;
                File.WriteAllText("out.csv", rescsv, Encoding.UTF8);
            }
        }
    }
}
