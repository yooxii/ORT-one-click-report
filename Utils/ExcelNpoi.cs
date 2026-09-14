using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;

namespace ORT一键报告.Utils
{
    /// <summary>
    /// NPOI 读写辅助：把原 EPPlus 里顺手就有的写法（按 1 基行列取单元格、整块区域套样式、
    /// "B12" 地址解析、插图定位尺寸）集中在这里，业务代码只调这些方法。
    /// 约定：本类对外行列参数一律 **1 基**（与 DataCell 一致），内部转成 NPOI 的 0 基。
    /// </summary>
    public static class ExcelNpoi
    {
        /// <summary>OOXML 里 1 像素对应的 EMU（1 英寸 = 914400 EMU，96 DPI → 9525）</summary>
        public const int EmuPerPixel = 9525;

        /// <summary>单元格样式缓存：同一工作簿里相同描述的样式复用，避免超出样式数量上限</summary>
        private static readonly ConditionalWeakTable<IWorkbook, Dictionary<string, ICellStyle>> StyleCaches = new();

        /* ###############################  打开 / 保存  ################################ */

        /// <summary>
        /// 读取工作簿：一次性读入内存，避免模板被 Excel 占用或本进程锁定文件
        /// </summary>
        public static XSSFWorkbook OpenRead(string path)
        {
            return new XSSFWorkbook(new MemoryStream(File.ReadAllBytes(path)));
        }

        /// <summary>
        /// 读取工作簿（已有字节）
        /// </summary>
        public static XSSFWorkbook OpenRead(byte[] bytes)
        {
            return new XSSFWorkbook(new MemoryStream(bytes));
        }

        /// <summary>
        /// 新建工作簿
        /// </summary>
        public static XSSFWorkbook Create() => new XSSFWorkbook();

        /// <summary>
        /// 保存到文件（目录不存在时自动创建）
        /// </summary>
        public static void Save(IWorkbook workbook, string path)
        {
            string dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(dir))
            {
                Directory.CreateDirectory(dir);
            }
            using FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None);
            workbook.Write(fs, true);
        }

        /// <summary>
        /// 取工作表（按 0 基索引，越界返回 null）
        /// </summary>
        public static ISheet SheetAt(IWorkbook workbook, int index)
            => workbook == null || index < 0 || index >= workbook.NumberOfSheets ? null : workbook.GetSheetAt(index);

        /// <summary>
        /// 取工作表（按名称，不存在返回 null）
        /// </summary>
        public static ISheet SheetByName(IWorkbook workbook, string name)
            => workbook?.GetSheet(name);

        /* ###############################  行列取值（1 基）  ################################ */

        /// <summary>
        /// 行号（1 基）：空表返回 0
        /// </summary>
        public static int LastRow(ISheet sheet) => sheet == null ? 0 : sheet.LastRowNum + 1;

        /// <summary>
        /// 最大列号（1 基）：空表返回 0
        /// </summary>
        public static int LastColumn(ISheet sheet)
        {
            if (sheet == null)
            {
                return 0;
            }
            int max = 0;
            for (int r = sheet.FirstRowNum; r <= sheet.LastRowNum; r++)
            {
                IRow row = sheet.GetRow(r);
                if (row != null && row.LastCellNum > max)
                {
                    max = row.LastCellNum;
                }
            }
            return max;
        }

        /// <summary>
        /// 取行（1 基）；不存在则创建
        /// </summary>
        public static IRow Row(ISheet sheet, int row1)
        {
            IRow row = sheet.GetRow(row1 - 1);
            return row ?? sheet.CreateRow(row1 - 1);
        }

        /// <summary>
        /// 取已存在的行（1 基）；不存在返回 null
        /// </summary>
        public static IRow ExistingRow(ISheet sheet, int row1) => sheet?.GetRow(row1 - 1);

        /// <summary>
        /// 取单元格（1 基）；不存在则创建
        /// </summary>
        public static ICell Cell(ISheet sheet, int row1, int col1)
        {
            IRow row = Row(sheet, row1);
            ICell cell = row.GetCell(col1 - 1);
            return cell ?? row.CreateCell(col1 - 1);
        }

        /// <summary>
        /// 取已存在的单元格（1 基）；不存在返回 null
        /// </summary>
        public static ICell ExistingCell(ISheet sheet, int row1, int col1) => sheet?.GetRow(row1 - 1)?.GetCell(col1 - 1);

        /// <summary>
        /// 单元格显示文本（等价 EPPlus 的 cell.Text：日期/数字都是所见即所得）。
        /// 两个必须特殊处理的地方：
        /// 1) 公式单元格：NPOI 的 DataFormatter 单独使用时返回公式文本，这里改用文件里缓存的最近计算结果；
        /// 2) 日期单元格：NPOI 的 DataFormatter 不会去掉数字格式里的转义引号（如 d"月"m"日" → 1"月"9"日"），
        ///    统一输出 yyyy/M/d（带时间的格式补上时分），保证下游日期解析稳定。
        /// </summary>
        public static string CellText(ISheet sheet, int row1, int col1)
        {
            ICell cell = ExistingCell(sheet, row1, col1);
            if (cell == null)
            {
                return "";
            }
            try
            {
                if (cell.CellType == CellType.Formula)
                {
                    return CachedText(cell);
                }
                if (cell.CellType == CellType.Numeric && DateUtil.IsCellDateFormatted(cell))
                {
                    return DateText(cell);
                }
                return new DataFormatter().FormatCellValue(cell);
            }
            catch
            {
                return cell.ToString() ?? "";
            }
        }

        /// <summary>
        /// 公式单元格的显示文本（取缓存的计算结果）
        /// </summary>
        private static string CachedText(ICell cell)
        {
            switch (cell.CachedFormulaResultType)
            {
                case CellType.String:
                    return cell.StringCellValue ?? "";
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.Numeric:
                    return DateUtil.IsCellDateFormatted(cell)
                        ? DateText(cell)
                        : cell.NumericCellValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
                case CellType.Blank:
                    return "";
                default:
                    return new DataFormatter().FormatCellValue(cell);
            }
        }

        /// <summary>
        /// 日期显示文本：yyyy/M/d（含时分时补 HH:mm）
        /// </summary>
        private static string DateText(ICell cell)
        {
            DateTime? value = cell.DateCellValue;
            if (value == null)
            {
                return new DataFormatter().FormatCellValue(cell);
            }
            return value.Value.TimeOfDay == TimeSpan.Zero
                ? value.Value.ToString("yyyy/M/d")
                : value.Value.ToString("yyyy/M/d HH:mm");
        }

        /// <summary>
        /// 单元格原始值的字符串形式（等价 EPPlus 的 cell.Value?.ToString()）
        /// </summary>
        public static string CellValue(ISheet sheet, int row1, int col1)
        {
            ICell cell = ExistingCell(sheet, row1, col1);
            if (cell == null)
            {
                return null;
            }
            switch (cell.CellType)
            {
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Numeric:
                    return DateUtil.IsCellDateFormatted(cell)
                        ? cell.DateCellValue?.ToString()
                        : cell.NumericCellValue.ToString();
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.Formula:
                    try
                    {
                        return cell.CachedFormulaResultType == CellType.String
                            ? cell.StringCellValue
                            : cell.NumericCellValue.ToString();
                    }
                    catch
                    {
                        return cell.ToString();
                    }
                default:
                    return cell.ToString();
            }
        }

        /// <summary>
        /// 写文本
        /// </summary>
        public static void SetCell(ISheet sheet, int row1, int col1, string text)
        {
            ICell cell = Cell(sheet, row1, col1);
            if (text == null)
            {
                cell.SetCellValue((string)null);
            }
            else
            {
                cell.SetCellValue(text);
            }
        }

        /// <summary>
        /// 写数字
        /// </summary>
        public static void SetCell(ISheet sheet, int row1, int col1, double value) => Cell(sheet, row1, col1).SetCellValue(value);

        /// <summary>
        /// 写任意值（等价 EPPlus 的 cell.Value = obj）：数字/日期按原类型写，其余转文本
        /// </summary>
        public static void SetCell(ISheet sheet, int row1, int col1, object value)
        {
            switch (value)
            {
                case null:
                    SetCell(sheet, row1, col1, (string)null);
                    break;
                case double d:
                    SetCell(sheet, row1, col1, d);
                    break;
                case int i:
                    SetCell(sheet, row1, col1, (double)i);
                    break;
                case DateTime dt:
                    SetCell(sheet, row1, col1, (DateTime?)dt);
                    break;
                default:
                    SetCell(sheet, row1, col1, value.ToString());
                    break;
            }
        }

        /// <summary>
        /// 写日期（带 yyyy/M/d 显示格式）
        /// </summary>
        public static void SetCell(ISheet sheet, int row1, int col1, DateTime? value)
        {
            ICell cell = Cell(sheet, row1, col1);
            if (value == null)
            {
                cell.SetCellValue((string)null);
                return;
            }
            cell.SetCellValue(value.Value);
            cell.CellStyle = DateStyleOf(sheet.Workbook);
        }

        /// <summary>
        /// 日期显示样式（同一工作簿复用）
        /// </summary>
        public static ICellStyle DateStyleOf(IWorkbook workbook)
            => Style(workbook, "date:yyyy/M/d", style => style.DataFormat = workbook.CreateDataFormat().GetFormat("yyyy/M/d"));

        /// <summary>
        /// 写公式
        /// </summary>
        public static void SetFormula(ISheet sheet, int row1, int col1, string formula) => Cell(sheet, row1, col1).CellFormula = formula;

        /// <summary>
        /// 给区域内每个单元格写同一个值（等价 EPPlus 的 ws.Cells[r1,c1,r2,c2].Value = x）
        /// </summary>
        public static void FillRange(ISheet sheet, int row1, int col1, int row2, int col2, object value)
        {
            int r1 = Math.Min(row1, row2), r2 = Math.Max(row1, row2);
            int c1 = Math.Min(col1, col2), c2 = Math.Max(col1, col2);
            for (int r = r1; r <= r2; r++)
            {
                for (int c = c1; c <= c2; c++)
                {
                    SetCell(sheet, r, c, value);
                }
            }
        }

        /// <summary>
        /// 合并单元格
        /// </summary>
        public static void Merge(ISheet sheet, int row1, int col1, int row2, int col2)
            => sheet.AddMergedRegion(new CellRangeAddress(row1 - 1, row2 - 1, col1 - 1, col2 - 1));

        /* ###############################  地址（1 基）  ################################ */

        /// <summary>
        /// 地址 → 行号（1 基，解析失败返回 0）
        /// </summary>
        public static int RowOf(string address)
        {
            try
            {
                return new CellReference(address)?.Row + 1 ?? 0;
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// 地址 → 列号（1 基，解析失败返回 0）
        /// </summary>
        public static int ColumnOf(string address)
        {
            try
            {
                return new CellReference(address)?.Col + 1 ?? 0;
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// 行列（1 基）→ 地址，如 (12, 2) → "B12"
        /// </summary>
        public static string AddressOf(int row1, int col1)
        {
            if (row1 <= 0 || col1 <= 0)
            {
                return "";
            }
            return new CellReference(row1 - 1, col1 - 1).FormatAsString();
        }

        /* ###############################  样式（整块套用）  ################################ */

        private static Dictionary<string, ICellStyle> CacheOf(IWorkbook workbook)
        {
            if (!StyleCaches.TryGetValue(workbook, out Dictionary<string, ICellStyle> cache))
            {
                cache = new Dictionary<string, ICellStyle>();
                StyleCaches.Add(workbook, cache);
            }
            return cache;
        }

        /// <summary>
        /// 取（或建）一种样式；key 相同则复用，避免超出 Excel 的样式数量限制
        /// </summary>
        private static ICellStyle Style(IWorkbook workbook, string key, Action<ICellStyle> configure)
        {
            Dictionary<string, ICellStyle> cache = CacheOf(workbook);
            if (!cache.TryGetValue(key, out ICellStyle style))
            {
                style = workbook.CreateCellStyle();
                configure(style);
                cache[key] = style;
            }
            return style;
        }

        /// <summary>
        /// 对区域套用现有样式（等价 EPPlus 的 ws.Cells[r1,c1,r2,c2].Style = style）
        /// </summary>
        public static void ApplyStyle(ISheet sheet, int row1, int col1, int row2, int col2, ICellStyle style)
        {
            if (sheet == null || style == null)
            {
                return;
            }
            if (row2 < row1)
            {
                (row1, row2) = (row2, row1);
            }
            if (col2 < col1)
            {
                (col1, col2) = (col2, col1);
            }
            for (int r = row1; r <= row2; r++)
            {
                for (int c = col1; c <= col2; c++)
                {
                    Cell(sheet, r, c).CellStyle = style;
                }
            }
        }

        /// <summary>
        /// 创建四边细边框样式
        /// </summary>
        public static ICellStyle BorderStyleOf(IWorkbook workbook, BorderStyle border = BorderStyle.Thin)
        {
            string key = "border:" + border;
            return Style(workbook, key, style =>
            {
                style.BorderTop = border;
                style.BorderBottom = border;
                style.BorderLeft = border;
                style.BorderRight = border;
            });
        }

        /// <summary>
        /// 创建外框加粗、内部细线的样式族（用于 EPPlus 的 BorderAround）：返回 [内部, 外框]
        /// </summary>
        public static (ICellStyle Inner, ICellStyle Outer) BorderAroundStyles(IWorkbook workbook, BorderStyle inner, BorderStyle outer)
        {
            ICellStyle innerStyle = Style(workbook, "border:" + inner, style =>
            {
                style.BorderTop = inner;
                style.BorderBottom = inner;
                style.BorderLeft = inner;
                style.BorderRight = inner;
            });
            ICellStyle outerStyle = Style(workbook, "borderAround:" + inner + ":" + outer, style =>
            {
                style.BorderTop = outer;
                style.BorderBottom = outer;
                style.BorderLeft = outer;
                style.BorderRight = outer;
            });
            return (innerStyle, outerStyle);
        }

        /// <summary>
        /// 创建填充样式
        /// </summary>
        public static ICellStyle FillStyleOf(IWorkbook workbook, short colorIndex, FillPattern pattern = FillPattern.SolidForeground)
        {
            string key = "fill:" + colorIndex + ":" + pattern;
            return Style(workbook, key, style =>
            {
                style.FillForegroundColor = colorIndex;
                style.FillPattern = pattern;
            });
        }

        /// <summary>
        /// 创建对齐样式（可叠加边框）
        /// </summary>
        public static ICellStyle AlignmentStyleOf(IWorkbook workbook, HorizontalAlignment horizontal, VerticalAlignment vertical)
        {
            string key = "align:" + horizontal + ":" + vertical;
            return Style(workbook, key, style =>
            {
                style.Alignment = horizontal;
                style.VerticalAlignment = vertical;
            });
        }

        /// <summary>
        /// 把样式的对齐方式设为居中（在同一 workbook 内复用）
        /// </summary>
        public static ICellStyle CenteredStyleOf(IWorkbook workbook)
            => AlignmentStyleOf(workbook, HorizontalAlignment.Center, VerticalAlignment.Center);

        /// <summary>
        /// 设置列宽（字符宽，等价 EPPlus 的 Column(i).Width）
        /// </summary>
        public static void SetColumnWidth(ISheet sheet, int col1, double characters)
            => sheet.SetColumnWidth(col1 - 1, (int)Math.Round(characters * 256));

        /// <summary>
        /// 设置行高（磅，等价 EPPlus 的 Row(i).Height，传入 &lt;=0 表示自动）
        /// </summary>
        public static void SetRowHeight(ISheet sheet, int row1, double points)
        {
            IRow row = Row(sheet, row1);
            if (points > 0)
            {
                row.HeightInPoints = (float)points;
            }
            else
            {
                row.Height = -1;
            }
        }

        /// <summary>
        /// 区域样式描述：等价 EPPlus 的 ws.Cells[r1,c1,r2,c2].Style.XXX 链式设置
        /// </summary>
        public sealed class CellStyleSpec
        {
            /// <summary>数字格式，如 "0.00"</summary>
            public string NumberFormat { get; set; }

            /// <summary>四边边框</summary>
            public BorderStyle? Border { get; set; }

            /// <summary>整块区域的外边框（只画在区域边上）</summary>
            public BorderStyle? OuterBorder { get; set; }

            public HorizontalAlignment? Horizontal { get; set; }
            public VerticalAlignment? Vertical { get; set; }
            public bool WrapText { get; set; }
            public double? FontSize { get; set; }

            /// <summary>填充色（RGB），为空则不填充</summary>
            public byte[] FillRgb { get; set; }
        }

        /// <summary>
        /// 对区域套用组合样式（内部颜色/对齐/格式，外框单独处理）
        /// </summary>
        public static void ApplyStyle(ISheet sheet, int row1, int col1, int row2, int col2, CellStyleSpec spec)
        {
            if (sheet == null || spec == null)
            {
                return;
            }
            int r1 = Math.Min(row1, row2), r2 = Math.Max(row1, row2);
            int c1 = Math.Min(col1, col2), c2 = Math.Max(col1, col2);
            for (int r = r1; r <= r2; r++)
            {
                for (int c = c1; c <= c2; c++)
                {
                    bool top = r == r1, bottom = r == r2, left = c == c1, right = c == c2;
                    bool useOuter = spec.OuterBorder.HasValue && (top || bottom || left || right);
                    Cell(sheet, r, c).CellStyle = BuildStyle(sheet.Workbook, spec, useOuter && top, useOuter && bottom, useOuter && left, useOuter && right);
                }
            }
        }

        private static ICellStyle BuildStyle(IWorkbook workbook, CellStyleSpec spec, bool top, bool bottom, bool left, bool right)
        {
            string fill = spec.FillRgb == null ? "" : string.Join(",", spec.FillRgb);
            string key = $"spec:{spec.NumberFormat}|{spec.Border}|{spec.OuterBorder}|{spec.Horizontal}|{spec.Vertical}|{spec.WrapText}|{spec.FontSize}|{fill}|{top}{bottom}{left}{right}";
            return Style(workbook, key, style =>
            {
                if (!string.IsNullOrEmpty(spec.NumberFormat))
                {
                    style.DataFormat = workbook.CreateDataFormat().GetFormat(spec.NumberFormat);
                }
                if (spec.Border is BorderStyle border)
                {
                    style.BorderTop = border;
                    style.BorderBottom = border;
                    style.BorderLeft = border;
                    style.BorderRight = border;
                }
                if (spec.OuterBorder is BorderStyle outer)
                {
                    if (top) style.BorderTop = outer;
                    if (bottom) style.BorderBottom = outer;
                    if (left) style.BorderLeft = outer;
                    if (right) style.BorderRight = outer;
                }
                if (spec.Horizontal is HorizontalAlignment horizontal)
                {
                    style.Alignment = horizontal;
                }
                if (spec.Vertical is VerticalAlignment vertical)
                {
                    style.VerticalAlignment = vertical;
                }
                if (spec.WrapText)
                {
                    style.WrapText = true;
                }
                if (spec.FontSize is double fontSize)
                {
                    IFont font = workbook.CreateFont();
                    font.FontHeightInPoints = (short)Math.Round(fontSize);
                    style.SetFont(font);
                }
                if (spec.FillRgb != null)
                {
                    if (style is XSSFCellStyle xssfStyle && spec.FillRgb.Length >= 3)
                    {
                        XSSFColor color = new XSSFColor();
                        color.SetRgb(spec.FillRgb);
                        xssfStyle.SetFillForegroundColor(color);
                    }
                    style.FillPattern = FillPattern.SolidForeground;
                }
            });
        }

        /// <summary>
        /// 单元格所在合并区域的标识（未合并且无合并区时返回 -1），用于判断相邻行是否属于同一合并块
        /// </summary>
        public static int MergeRegionId(ISheet sheet, int row1, int col1)
        {
            if (sheet == null)
            {
                return -1;
            }
            int index = 0;
            foreach (CellRangeAddress region in sheet.MergedRegions)
            {
                if (region.IsInRange(row1 - 1, col1 - 1))
                {
                    return index;
                }
                index++;
            }
            return -1;
        }

        /* ###############################  行操作（插入/删除/复制样式）  ################################ */

        /// <summary>
        /// 行高（磅），未设置返回 -1
        /// </summary>
        public static double RowHeight(ISheet sheet, int row1)
        {
            IRow row = ExistingRow(sheet, row1);
            return row == null ? -1 : row.HeightInPoints;
        }

        /// <summary>
        /// 复制整行样式（含单元格样式，不含值）
        /// </summary>
        public static void CopyRowStyle(ISheet sheet, int srcRow1, int dstRow1, int columnCount = -1)
        {
            IRow src = ExistingRow(sheet, srcRow1);
            if (src == null)
            {
                return;
            }
            IRow dst = Row(sheet, dstRow1);
            int last = columnCount > 0 ? columnCount : Math.Max(LastColumn(sheet), src.LastCellNum);
            for (int c = 1; c <= last; c++)
            {
                ICell srcCell = src.GetCell(c - 1);
                if (srcCell == null)
                {
                    continue;
                }
                ICell dstCell = dst.GetCell(c - 1) ?? dst.CreateCell(c - 1);
                dstCell.CellStyle = srcCell.CellStyle;
            }
            dst.HeightInPoints = src.HeightInPoints;
        }

        /// <summary>
        /// 把某列的样式与公式复制到另一列（等价 EPPlus 的 CopyStyles + CopyFormulas）
        /// </summary>
        public static void CopyColumnStylesAndFormulas(ISheet sheet, int srcCol1, int dstCol1, int row1, int row2)
        {
            for (int r = row1; r <= row2; r++)
            {
                ICell src = ExistingCell(sheet, r, srcCol1);
                if (src == null)
                {
                    continue;
                }
                ICell dst = Cell(sheet, r, dstCol1);
                dst.CellStyle = src.CellStyle;
                if (src.CellType == CellType.Formula && !string.IsNullOrEmpty(src.CellFormula))
                {
                    dst.CellFormula = src.CellFormula;
                }
            }
        }

        /// <summary>
        /// 在某行之前插入若干空行（1 基，等价 EPPlus 的 InsertRow）
        /// </summary>
        public static void InsertRows(ISheet sheet, int row1, int count)
        {
            if (count <= 0)
            {
                return;
            }
            int start = row1 - 1;
            int last = sheet.LastRowNum;
            if (start <= last)
            {
                sheet.ShiftRows(start, last, count, true, false);
            }
            for (int i = 0; i < count; i++)
            {
                if (sheet.GetRow(start + i) == null)
                {
                    sheet.CreateRow(start + i);
                }
            }
        }

        /// <summary>
        /// 从某行起删除若干行（1 基，等价 EPPlus 的 DeleteRow）
        /// </summary>
        public static void DeleteRows(ISheet sheet, int row1, int count)
        {
            if (count <= 0)
            {
                return;
            }
            int start = row1 - 1;
            int last = sheet.LastRowNum;
            // 先上移后续行，再删掉尾部残留行（POI 推荐顺序）
            if (start + count <= last)
            {
                sheet.ShiftRows(start + count, last, -count, true, false);
            }
            for (int i = 0; i < count; i++)
            {
                IRow row = sheet.GetRow(last - i);
                if (row != null)
                {
                    sheet.RemoveRow(row);
                }
            }
        }

        /* ###############################  图片  ################################ */

        /// <summary>
        /// 插入图片：左上角锚在 (row1, col1)，按像素指定宽高
        /// </summary>
        public static IPicture AddPicture(IWorkbook workbook, ISheet sheet, byte[] imageBytes, PictureType type,
            int row1, int col1, int widthPx = 300, int heightPx = 220, int offsetXPx = 0, int offsetYPx = 0)
        {
            if (imageBytes == null || imageBytes.Length == 0)
            {
                return null;
            }
            int pictureIndex = workbook.AddPicture(imageBytes, type);
            IDrawing drawing = sheet.CreateDrawingPatriarch();
            IClientAnchor anchor = workbook.GetCreationHelper().CreateClientAnchor();
            anchor.Col1 = col1 - 1;
            anchor.Row1 = row1 - 1;
            anchor.Dx1 = offsetXPx * EmuPerPixel;
            anchor.Dy1 = offsetYPx * EmuPerPixel;
            // 尺寸用 EMU 偏移表达：锚点仍落在同一单元格，终点用偏移撑开
            anchor.Col2 = col1 - 1;
            anchor.Row2 = row1 - 1;
            anchor.Dx2 = (offsetXPx + widthPx) * EmuPerPixel;
            anchor.Dy2 = (offsetYPx + heightPx) * EmuPerPixel;
            return drawing.CreatePicture(anchor, pictureIndex);
        }

        /// <summary>
        /// 读取工作表中的图片（左上角 1 基行列 + 字节）
        /// </summary>
        public static List<(int Row, int Column, string Name, byte[] Bytes)> Pictures(ISheet sheet)
        {
            List<(int, int, string, byte[])> list = [];
            if (sheet == null)
            {
                return list;
            }
            if (sheet.CreateDrawingPatriarch() is not XSSFDrawing drawing)
            {
                return list;
            }
            foreach (XSSFShape shape in drawing.GetShapes())
            {
                if (shape is XSSFPicture picture)
                {
                    IClientAnchor anchor = picture.ClientAnchor;
                    byte[] bytes = picture.PictureData?.Data;
                    list.Add((anchor.Row1 + 1, anchor.Col1 + 1, picture.Name, bytes));
                }
            }
            return list;
        }
    }
}
