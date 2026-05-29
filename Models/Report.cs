using CommunityToolkit.Mvvm.ComponentModel;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.Collections.Generic;
using System.Windows.Media;

namespace ORT一键报告.Models
{

    public class DataCell : ObservableObject
    {
        private string _data;
        public string Data { get => _data; set => SetProperty(ref _data, value); }

        private int _row;
        public int Row { get => _row; set => SetProperty(ref _row, value); }

        private int _column;
        public int Column { get => _column; set => SetProperty(ref _column, value); }

        public List<ExcelPictureInfo> Images { get; set; } = null;
        public string TopLeftAddress
        {
            get => ExcelCellBase.GetAddress(Row, Column);
            set
            {
                int bRow = Row;
                int bColumn = Column;
                try
                {
                    ExcelAddress Addr = new(value);
                    Row = Addr.Start.Row;
                    Column = Addr.Start.Column;
                }
                catch
                {
                    Row = bRow;
                    Column = bColumn;
                }
            }
        }
        public override string ToString()
        {
            return $"{Data} - {TopLeftAddress}({Row},{Column})";
        }
        public bool IsNotEmpty()
        {
            return !string.IsNullOrWhiteSpace(_data);
        }
        public static bool IsNotNull(DataCell cell)
        {
            return cell is not null && !string.IsNullOrEmpty(cell._data);
        }
    }

    public class TestItemInfo
    {
        public string TestItemName { get; set; }
        public string Date { get; set; }
    }

    public class UUTInfoFromExcel
    {
        public List<string> SNs { get; set; }
        public string WorkOrder { get; set; }
        public string Revision { get; set; }
        public string DC { get; set; }
        public List<TestItemInfo> TestItems { get; set; }

        public override string ToString()
        {
            return $"{WorkOrder},{Revision},{DC},{(TestItems == null ? 0 : TestItems.Count)},{(SNs == null ? 0 : SNs.Count)}";
        }
    }

    public class ResultDetails
    {
        public string BIroom { get; set; } = "";
        public string BIarea { get; set; } = "";
        public string BIplace { get; set; } = "";
        public string SN { get; set; } = "";
        public string WorkOrder { get; set; } = "";
        public string Version { get; set; } = "";
        public string DC { get; set; } = "";
        public ReportStatus InspectionPrev { get; set; }
        public ReportStatus FunPrev { get; set; }
        public ReportStatus InspectionAfter { get; set; }
        public ReportStatus FunAfter { get; set; }
        public ReportStatus HiPot { get; set; }
        public string Comments { get; set; } = "";
    }

    /// <summary>
    /// 辅助类：用于返回提取的图片信息
    /// </summary>
    public class ExcelPictureInfo
    {
        public ExcelPicture Picture { get; set; } // 原始对象
        public ImageSource ImageSrc { get; set; }    // System.Drawing.Image 对象
        public byte[] ImageBytes { get; set; }    // 字节数组
        public string Name { get; set; }          // 图片名称
    }


    public class EMIUUTdataInfo
    {
        public List<string> SN { get; set; } = new List<string>();
        public List<string> Voltage { get; set; } = new List<string>();
        public List<string> Load { get; set; } = new List<string>();
        public List<string> LISN { get; set; } = new List<string>();

        /// <summary>
        /// 返回该机种的信息，字符串形式
        /// </summary>
        /// <param name="n">控制返回信息种类的个数，最大3</param>
        /// <returns></returns>
        public string GetName(int n)
        {
            string res = "";
            if (n >= 1)
                res += $"-{string.Join("_", Voltage.ToArray())}";
            if (n >= 2)
                res += $"-{string.Join("_", Load.ToArray())}".Replace("%", "");
            if (n >= 3)
                res += $"-{string.Join("_", LISN.ToArray())}";
            return res;
        }
    }

    public class EMIUUTData
    {
        public string Name => $"{SN}-{Voltage}-{Load}-{LISN}";
        public string SN { get; set; }
        public string Voltage { get; set; }
        public string Load { get; set; }
        public string LISN { get; set; }
        public string Model { get; set; }

        public List<List<float>> Datas { get; set; }
        public List<float> MinDatas { get; set; }

        public EMIUUTData()
        {
            Datas = [];
            MinDatas = [];
        }
    }
}
