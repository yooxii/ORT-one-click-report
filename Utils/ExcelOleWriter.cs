using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace ORT一键报告.Utils
{
    /// <summary>
    /// OLE 附件直写器：不依赖 Excel，直接把嵌入对象写进已保存的 xlsx 包。
    ///
    /// 用途：Excel COM 不可用或无法保存时的兜底（原 EPPlus 版本同样是"不依赖 Excel"就能嵌附件）。
    /// 写入结构与 Excel 自身产出一致：
    ///   xl/embeddings/&lt;文件名&gt;              嵌入的附件本体
    ///   xl/media/imageN.emf                  附件图标（VML 形状引用的图片）
    ///   xl/drawings/vmlDrawingN.vml(+rels)    该工作表全部图标形状与 ClientData（含锚点）
    ///   xl/worksheets/sheetN.xml              追加 &lt;legacyDrawing/&gt; 与 &lt;oleObjects&gt;（含 x14 锚点）
    ///   xl/worksheets/_rels/sheetN.xml.rels   追加 oleObject / vmlDrawing 关系
    ///   [Content_Types].xml                   补充新扩展名的默认内容类型
    /// </summary>
    public static class ExcelOleWriter
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace PkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly XNamespace Ct = "http://schemas.openxmlformats.org/package/2006/content-types";
        private static readonly XNamespace Xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        private const int EmuPerPixel = 9525;
        private const double PointsPerPixel = 72.0 / 96.0;

        private const string OleRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject";
        private const string PackageRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package";
        private const string VmlRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
        private const string ImageRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
        private const string OlePackageContentType = "application/vnd.openxmlformats-officedocument.oleObject";

        /// <summary>
        /// 把 OLE 对象写进 xlsx（就地覆盖保存），返回成功写入的个数
        /// </summary>
        public static int Write(string xlsxPath, IReadOnlyList<OleEmbedRequest> requests)
        {
            List<OleEmbedRequest> items = requests?.Where(r => r != null && !string.IsNullOrWhiteSpace(r.ObjectPath) && File.Exists(r.ObjectPath)).ToList() ?? [];
            if (items.Count == 0)
            {
                return 0;
            }

            Dictionary<string, byte[]> parts = ReadParts(xlsxPath);
            if (!parts.TryGetValue("xl/workbook.xml", out byte[] workbookBytes))
            {
                _logger.Warn($"OLE 直写跳过：包内缺少 xl/workbook.xml（{Path.GetFileName(xlsxPath)}）");
                return 0;
            }

            Dictionary<string, string> sheetParts = ResolveSheetParts(parts, workbookBytes);
            Dictionary<string, SheetEdits> edits = new(StringComparer.OrdinalIgnoreCase);
            XDocument contentTypes = parts.TryGetValue("[Content_Types].xml", out byte[] ctBytes)
                ? ParsePart(ctBytes)
                : new XDocument(new XElement(Ct + "Types"));
            HashSet<string> addedExtensions = new(StringComparer.OrdinalIgnoreCase);

            int written = 0;
            int mediaIndex = NextIndex(parts.Keys, "xl/media/image", ".emf", ".png", ".jpg", ".jpeg");
            int vmlIndex = NextIndex(parts.Keys, "xl/drawings/vmlDrawing", ".vml");
            int embedIndex = 1;

            foreach (OleEmbedRequest req in items)
            {
                string sheetName = string.IsNullOrWhiteSpace(req.SheetName) ? null : req.SheetName;
                string sheetPart = sheetName != null && sheetParts.TryGetValue(sheetName, out string p)
                    ? p
                    : sheetParts.Values.FirstOrDefault();
                int row1 = ExcelNpoi.RowOf(req.TopLeftAddress);
                int col1 = ExcelNpoi.ColumnOf(req.TopLeftAddress);
                if (sheetPart == null || row1 <= 0 || col1 <= 0)
                {
                    _logger.Warn($"OLE 直写跳过（工作表/地址无效 {req.SheetName}/{req.TopLeftAddress}）：{Path.GetFileName(req.ObjectPath)}");
                    continue;
                }

                if (!edits.TryGetValue(sheetPart, out SheetEdits edit))
                {
                    edit = new SheetEdits(sheetPart);
                    edits[sheetPart] = edit;
                }

                // 1) 附件本体
                string ext = Path.GetExtension(req.ObjectPath).ToLowerInvariant();
                string embedPart = MakeUnique(parts, "xl/embeddings/" + SanitizeFileName(Path.GetFileName(req.ObjectPath)), embedIndex++);
                parts[embedPart] = File.ReadAllBytes(req.ObjectPath);
                EnsureContentType(contentTypes, addedExtensions, ext);

                string embedRelId = edit.NextRelId();
                edit.Relationships.Add(new XElement(PkgRel + "Relationship",
                    new XAttribute("Id", embedRelId),
                    new XAttribute("Type", IsPackageTarget(ext) ? PackageRelType : OleRelType),
                    new XAttribute("Target", "../embeddings/" + Path.GetFileName(embedPart))));

                // 2) 图标（该工作表共用一个 vmlDrawing 部件）
                int shapeId = 1025 + written;
                byte[] iconBytes = LoadIconBytes(req.IconPath);
                bool hasIcon = iconBytes != null && iconBytes.Length > 0;
                if (hasIcon)
                {
                    if (edit.VmlPart == null)
                    {
                        edit.VmlPart = "xl/drawings/vmlDrawing" + vmlIndex++ + ".vml";
                        // VML 部件必须声明内容类型，否则整个包非法（Excel/NPOI 都打不开）
                        EnsureContentType(contentTypes, addedExtensions, ".vml");
                        edit.LegacyRelId = edit.NextRelId();
                        edit.Relationships.Add(new XElement(PkgRel + "Relationship",
                            new XAttribute("Id", edit.LegacyRelId),
                            new XAttribute("Type", VmlRelType),
                            new XAttribute("Target", "../drawings/" + Path.GetFileName(edit.VmlPart))));
                    }
                    string mediaPart = "xl/media/image" + mediaIndex++ + ".emf";
                    parts[mediaPart] = iconBytes;
                    EnsureContentType(contentTypes, addedExtensions, ".emf");

                    string mediaRelId = "rId" + (edit.VmlShapes.Count + 1);
                    edit.VmlShapeRelIds.Add(mediaRelId);
                    edit.VmlShapes.Add(BuildVmlShape(shapeId, row1, col1, req, mediaRelId));
                    edit.VmlMediaNames.Add(Path.GetFileName(mediaPart));
                }

                // 3) 工作表里的 oleObject
                edit.OleXml.Add(BuildOleObject(shapeId, row1, col1, req, embedRelId, hasIcon ? edit.LegacyRelId : null, hasIcon));
                written++;
            }

            if (written == 0)
            {
                return 0;
            }

            // VML 部件与关系
            foreach (SheetEdits edit in edits.Values)
            {
                if (edit.VmlPart == null)
                {
                    continue;
                }
                parts[edit.VmlPart] = Encoding.UTF8.GetBytes(BuildVml(edit));
                parts[VmlRelsPath(edit.VmlPart)] = Encoding.UTF8.GetBytes(BuildVmlRels(edit));
            }

            // 回写工作表：追加 legacyDrawing + oleObjects（放在 worksheet 末尾，符合 schema 顺序）
            foreach (SheetEdits edit in edits.Values)
            {
                if (!parts.TryGetValue(edit.SheetPart, out byte[] sheetBytes))
                {
                    continue;
                }
                XDocument sheet = ParsePart(sheetBytes);
                XElement root = sheet.Root;
                if (root == null)
                {
                    continue;
                }
                if (edit.LegacyRelId != null && root.Element(Main + "legacyDrawing") == null)
                {
                    root.Add(new XElement(Main + "legacyDrawing", new XAttribute(Rel + "id", edit.LegacyRelId)));
                }
                XElement oleObjects = root.Element(Main + "oleObjects");
                if (oleObjects == null)
                {
                    oleObjects = new XElement(Main + "oleObjects");
                    root.Add(oleObjects);
                }
                foreach (XElement ole in edit.OleXml)
                {
                    oleObjects.Add(ole);
                }
                parts[edit.SheetPart] = Encoding.UTF8.GetBytes(sheet.ToString(SaveOptions.DisableFormatting));

                string relsPath = SheetRelsPath(edit.SheetPart);
                XDocument rels = parts.TryGetValue(relsPath, out byte[] relBytes)
                    ? ParsePart(relBytes)
                    : new XDocument(new XElement(PkgRel + "Relationships"));
                foreach (XElement relationship in edit.Relationships)
                {
                    rels.Root.Add(relationship);
                }
                parts[relsPath] = Encoding.UTF8.GetBytes(rels.ToString(SaveOptions.DisableFormatting));
            }

            parts["[Content_Types].xml"] = Encoding.UTF8.GetBytes(contentTypes.ToString(SaveOptions.DisableFormatting));

            WriteParts(xlsxPath, parts);
            _logger.Info($"OLE 直写完成：{written}/{items.Count} 个附件写入 {Path.GetFileName(xlsxPath)}");
            return written;
        }

        /// <summary>
        /// 单个工作表要追加的内容
        /// </summary>
        private sealed class SheetEdits
        {
            private int _relSeq;

            public SheetEdits(string sheetPart)
            {
                SheetPart = sheetPart;
            }

            public string SheetPart { get; }
            public List<XElement> Relationships { get; } = [];
            public List<XElement> OleXml { get; } = [];
            public string VmlPart { get; set; }
            public string LegacyRelId { get; set; }
            public List<string> VmlShapes { get; } = [];
            public List<string> VmlShapeRelIds { get; } = [];
            public List<string> VmlMediaNames { get; } = [];

            public string NextRelId() => "rIdOle" + (++_relSeq);
        }

        /* ###############################  包读写  ################################ */

        /// <summary>
        /// 解析 XML 部件：按字节流读，自动处理 BOM 与声明（直接 GetString 后 Parse 会因 BOM 报"根级别上的数据无效"）
        /// </summary>
        private static XDocument ParsePart(byte[] bytes)
            => XDocument.Load(new MemoryStream(bytes), LoadOptions.PreserveWhitespace);

        private static Dictionary<string, byte[]> ReadParts(string xlsxPath)
        {
            Dictionary<string, byte[]> parts = new(StringComparer.OrdinalIgnoreCase);
            using ZipArchive zip = ZipFile.OpenRead(xlsxPath);
            foreach (ZipArchiveEntry entry in zip.Entries)
            {
                if (entry.FullName.EndsWith("/"))
                {
                    continue;
                }
                using Stream s = entry.Open();
                using MemoryStream ms = new();
                s.CopyTo(ms);
                parts[entry.FullName] = ms.ToArray();
            }
            return parts;
        }

        private static void WriteParts(string xlsxPath, Dictionary<string, byte[]> parts)
        {
            string tempPath = xlsxPath + ".ole.tmp";
            using (FileStream fs = new(tempPath, FileMode.Create, FileAccess.Write, FileShare.None))
            using (ZipArchive zip = new(fs, ZipArchiveMode.Create))
            {
                // [Content_Types].xml 放最前
                foreach (string key in parts.Keys.OrderBy(k => k == "[Content_Types].xml" ? 0 : 1))
                {
                    ZipArchiveEntry entry = zip.CreateEntry(key, CompressionLevel.Optimal);
                    using Stream s = entry.Open();
                    s.Write(parts[key], 0, parts[key].Length);
                }
            }
            File.Delete(xlsxPath);
            File.Move(tempPath, xlsxPath);
        }

        private static Dictionary<string, string> ResolveSheetParts(Dictionary<string, byte[]> parts, byte[] workbookBytes)
        {
            Dictionary<string, string> result = new(StringComparer.OrdinalIgnoreCase);
            XDocument workbook = ParsePart(workbookBytes);
            if (!parts.TryGetValue("xl/_rels/workbook.xml.rels", out byte[] relBytes))
            {
                return result;
            }
            XDocument rels = ParsePart(relBytes);
            foreach (XElement sheet in workbook.Descendants(Main + "sheet"))
            {
                string name = (string)sheet.Attribute("name");
                string rId = (string)sheet.Attribute(Rel + "id");
                if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(rId))
                {
                    continue;
                }
                string target = (string)rels.Descendants(PkgRel + "Relationship")
                    .FirstOrDefault(r => (string)r.Attribute("Id") == rId)?.Attribute("Target");
                if (!string.IsNullOrEmpty(target))
                {
                    result[name] = ResolvePart("xl/workbook.xml", target);
                }
            }
            return result;
        }

        private static string ResolvePart(string sourcePart, string target)
        {
            int slash = sourcePart.LastIndexOf('/');
            string combined = (slash >= 0 ? sourcePart.Substring(0, slash) : "") + "/" + target;
            List<string> segments = [];
            foreach (string segment in combined.Split('/'))
            {
                if (segment == "..")
                {
                    if (segments.Count > 0)
                    {
                        segments.RemoveAt(segments.Count - 1);
                    }
                }
                else if (segment != "." && segment.Length > 0)
                {
                    segments.Add(segment);
                }
            }
            return string.Join("/", segments);
        }

        private static string SheetRelsPath(string sheetPart)
        {
            int slash = sheetPart.LastIndexOf('/');
            return sheetPart.Substring(0, slash) + "/_rels/" + sheetPart.Substring(slash + 1) + ".rels";
        }

        private static string VmlRelsPath(string vmlPart)
        {
            int slash = vmlPart.LastIndexOf('/');
            return vmlPart.Substring(0, slash) + "/_rels/" + vmlPart.Substring(slash + 1) + ".rels";
        }

        private static int NextIndex(IEnumerable<string> keys, string prefix, params string[] suffixes)
        {
            int max = 0;
            foreach (string key in keys)
            {
                if (!key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                string tail = key.Substring(prefix.Length);
                int end = tail.IndexOf('.');
                if (end > 0 && suffixes.Any(s => tail.Substring(end).Equals(s, StringComparison.OrdinalIgnoreCase))
                    && int.TryParse(tail.Substring(0, end), out int n) && n > max)
                {
                    max = n;
                }
            }
            return max + 1;
        }

        private static string MakeUnique(Dictionary<string, byte[]> parts, string path, int index)
        {
            if (!parts.ContainsKey(path))
            {
                return path;
            }
            string dir = path.Substring(0, path.LastIndexOf('/') + 1);
            string name = Path.GetFileNameWithoutExtension(path);
            string ext = Path.GetExtension(path);
            string candidate = $"{dir}{name}_{index}{ext}";
            while (parts.ContainsKey(candidate))
            {
                candidate = $"{dir}{name}_{++index}{ext}";
            }
            return candidate;
        }

        private static string SanitizeFileName(string name)
        {
            StringBuilder sb = new();
            foreach (char c in name)
            {
                sb.Append(char.IsLetterOrDigit(c) || c == '.' || c == '-' || c == '_' ? c : '_');
            }
            return sb.Length == 0 ? "oleObject.bin" : sb.ToString();
        }

        private static bool IsPackageTarget(string ext)
            => ext is ".xlsx" or ".xlsm" or ".xlsb" or ".docx" or ".docm" or ".pptx" or ".pptm";

        private static void EnsureContentType(XDocument contentTypes, HashSet<string> added, string ext)
        {
            string extension = string.IsNullOrEmpty(ext) ? "bin" : ext.TrimStart('.').ToLowerInvariant();
            if (!added.Add(extension))
            {
                return;
            }
            if (contentTypes.Root.Elements(Ct + "Default")
                .Any(d => string.Equals((string)d.Attribute("Extension"), extension, StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }
            contentTypes.Root.Add(new XElement(Ct + "Default",
                new XAttribute("Extension", extension),
                new XAttribute("ContentType", ContentTypeOf(extension))));
        }

        private static string ContentTypeOf(string extension) => extension switch
        {
            "xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "xlsm" => "application/vnd.ms-excel.sheet.macroEnabled.12",
            "xls" => "application/vnd.ms-excel",
            "xlsb" => "application/vnd.ms-excel.sheet.binary.macroEnabled.12",
            "docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "doc" => "application/msword",
            "pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            "pdf" => "application/pdf",
            "zip" => OlePackageContentType,
            "emf" => "image/x-emf",
            "png" => "image/png",
            "jpg" => "image/jpeg",
            "jpeg" => "image/jpeg",
            "gif" => "image/gif",
            "vml" => "application/vnd.openxmlformats-officedocument.vmlDrawing",
            _ => OlePackageContentType,
        };

        /* ###############################  生成 XML 片段  ################################ */

        private static byte[] LoadIconBytes(string iconPath)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(iconPath) && File.Exists(iconPath))
                {
                    return File.ReadAllBytes(iconPath);
                }
                return Resources.image_xlsx_emf;
            }
            catch (Exception ex)
            {
                _logger.Warn($"读取 OLE 图标失败（改用内置图标）：{ex.Message}");
                return Resources.image_xlsx_emf;
            }
        }

        /// <summary>
        /// 工作表里的 oleObject 片段（含 x14 锚点；带图标时附带 objectPr → vmlDrawing 关系）。
        /// 命名空间按 Excel 的实际写法：oleObject/objectPr/anchor/from/to 属于 spreadsheetml 主命名空间，
        /// 内部的 col/row 等属于 spreadsheetDrawing 命名空间。
        /// </summary>
        private static XElement BuildOleObject(int shapeId, int row1, int col1, OleEmbedRequest req, string embedRelId, string vmlRelId, bool hasIcon)
        {
            int colOff = req.OffsetXPx * EmuPerPixel;
            int rowOff = req.OffsetYPx * EmuPerPixel;
            int width = req.WidthPx > 0 ? req.WidthPx : 100;
            int height = req.HeightPx > 0 ? req.HeightPx : 100;

            XElement anchor = new(Main + "anchor",
                new XAttribute("moveWithCells", "1"),
                new XAttribute(XNamespace.Xmlns + "xdr", Xdr.NamespaceName),
                new XElement(Main + "from",
                    new XElement(Xdr + "col", col1 - 1),
                    new XElement(Xdr + "colOff", colOff),
                    new XElement(Xdr + "row", row1 - 1),
                    new XElement(Xdr + "rowOff", rowOff)),
                new XElement(Main + "to",
                    new XElement(Xdr + "col", col1 - 1),
                    new XElement(Xdr + "colOff", colOff + (width * EmuPerPixel)),
                    new XElement(Xdr + "row", row1 - 1),
                    new XElement(Xdr + "rowOff", rowOff + (height * EmuPerPixel))));

            XElement objectPr = new(Main + "objectPr",
                new XAttribute("defaultSize", "0"),
                new XAttribute("autoPict", "0"));
            if (hasIcon && vmlRelId != null)
            {
                objectPr.Add(new XAttribute(Rel + "id", vmlRelId));
            }
            objectPr.Add(anchor);

            XElement oleObject = new(Main + "oleObject",
                new XAttribute("progId", "Package"),
                new XAttribute("dvAspect", "DVASPECT_ICON"),
                new XAttribute("shapeId", shapeId),
                new XAttribute(Rel + "id", embedRelId),
                objectPr);

            return new XElement(Mc + "AlternateContent",
                new XAttribute(XNamespace.Xmlns + "mc", Mc.NamespaceName),
                new XElement(Mc + "Choice",
                    new XAttribute("Requires", "x14"),
                    oleObject),
                new XElement(Mc + "Fallback",
                    new XElement(Main + "oleObject",
                        new XAttribute("progId", "Package"),
                        new XAttribute("dvAspect", "DVASPECT_ICON"),
                        new XAttribute("shapeId", shapeId),
                        new XAttribute(Rel + "id", embedRelId))));
        }

        /// <summary>
        /// VML 图标形状（Excel 通过 legacyDrawing 显示 OLE 对象的图标）
        /// </summary>
        private static string BuildVmlShape(int shapeId, int row1, int col1, OleEmbedRequest req, string mediaRelId)
        {
            int width = req.WidthPx > 0 ? req.WidthPx : 100;
            int height = req.HeightPx > 0 ? req.HeightPx : 100;
            string style = $"position:absolute;margin-left:{col1 * 48 + req.OffsetXPx}pt;margin-top:{row1 * 15 + req.OffsetYPx}pt;"
                + $"width:{width * PointsPerPixel:0.##}pt;height:{height * PointsPerPixel:0.##}pt;z-index:1;mso-wrap-style:tight";
            string anchor = $"{col1 - 1}, {req.OffsetXPx}, {row1 - 1}, {req.OffsetYPx}, {col1 - 1}, {req.OffsetXPx + width}, {row1 - 1}, {req.OffsetYPx + height}";
            return $"<v:shape id=\"_x0000_s{shapeId}\" type=\"#_x0000_t75\" style='{style}' filled=\"t\" o:insetmode=\"auto\">"
                + "<v:fill color2=\"window [65]\"/>"
                + $"<v:imagedata o:relid=\"{mediaRelId}\" o:title=\"\"/>"
                // Excel 对"以图标显示的 OLE 对象"写的也是 ObjectType="Pict"（OLE 语义在工作表的 <oleObject> 上）；
                // 写成 "Embed" 会让 NPOI 解析 VML 时抛异常（ObjectType 被当作枚举，没有 Embed 值）
                + "<x:ClientData ObjectType=\"Pict\"><x:SizeWithCells/>"
                + $"<x:Anchor>{anchor}</x:Anchor>"
                + "<x:AutoFill>False</x:AutoFill><x:CF>Pict</x:CF></x:ClientData></v:shape>";
        }

        private static string BuildVml(SheetEdits edit)
        {
            StringBuilder sb = new();
            sb.Append("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">");
            sb.Append("<v:shapetype id=\"_x0000_t75\" coordsize=\"21600,21600\" o:spt=\"75\" o:preferrelative=\"t\" path=\"m@4@5l@4@11@9@11@9@5xe\" filled=\"f\" stroked=\"f\">");
            sb.Append("<v:stroke joinstyle=\"miter\"/><v:formulas>");
            sb.Append("<v:f eqn=\"if lineDrawn pixelLineWidth 0\"/><v:f eqn=\"sum @0 1 0\"/><v:f eqn=\"sum 0 0 @1\"/><v:f eqn=\"prod @2 1 2\"/>");
            sb.Append("<v:f eqn=\"prod @3 21600 pixelWidth\"/><v:f eqn=\"prod @3 21600 pixelHeight\"/><v:f eqn=\"sum @0 0 1\"/><v:f eqn=\"prod @6 1 2\"/>");
            sb.Append("<v:f eqn=\"prod @7 21600 pixelWidth\"/><v:f eqn=\"sum @8 21600 0\"/><v:f eqn=\"prod @7 21600 pixelHeight\"/><v:f eqn=\"sum @10 21600 0\"/>");
            sb.Append("</v:formulas><v:path o:extrusionok=\"f\" gradientshapeok=\"t\" o:connecttype=\"rect\"/><o:lock v:ext=\"edit\" aspectratio=\"t\"/></v:shapetype>");
            foreach (string shape in edit.VmlShapes)
            {
                sb.Append(shape);
            }
            sb.Append("</xml>");
            return sb.ToString();
        }

        private static string BuildVmlRels(SheetEdits edit)
        {
            StringBuilder sb = new();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            for (int i = 0; i < edit.VmlShapes.Count; i++)
            {
                sb.Append($"<Relationship Id=\"{edit.VmlShapeRelIds[i]}\" Type=\"{ImageRelType}\" Target=\"../media/{edit.VmlMediaNames[i]}\"/>");
            }
            sb.Append("</Relationships>");
            return sb.ToString();
        }
    }
}
