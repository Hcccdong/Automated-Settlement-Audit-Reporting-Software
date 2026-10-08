using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace SettlementReview {
    // Physical workbook reading only; review parsing stays in ExcelReview.
    public static class ExcelWorkbookReader {
        static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        static int[] Address(string address) { var m = Regex.Match(address, @"^([A-Z]+)(\d+)$"); if (!m.Success) throw new InvalidDataException("无法识别的 Excel 单元格地址。"); int col = 0; foreach (char c in m.Groups[1].Value) col = col * 26 + c - 'A' + 1; return new[] { int.Parse(m.Groups[2].Value), col }; }
        static XDocument Xml(ZipArchiveEntry entry) { if (entry == null || entry.Length > 64000000) throw new InvalidDataException("Excel 文件内容缺失或工作表过大。"); using (var stream = entry.Open()) return XDocument.Load(stream); }
        public static List<ExcelSheet> Read(string path) {
            if (Path.GetExtension(path).Equals(".xls", StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("旧式 .xls 请先在 WPS 或 Excel 中另存为 .xlsx，再导入。原文件无需修改。");
            using (var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) using (var zip = new ZipArchive(file, ZipArchiveMode.Read)) {
                if (zip.Entries.Sum(e => e.Length) > 160000000) throw new InvalidDataException("Excel 内容过大，请拆分后导入。");
                var stringsEntry = zip.GetEntry("xl/sharedStrings.xml");
                var strings = stringsEntry == null ? new List<string>() : Xml(stringsEntry).Root.Elements(S + "si").Select(e => string.Concat(e.Descendants(S + "t").Select(t => t.Value))).ToList();
                var relationships = Xml(zip.GetEntry("xl/_rels/workbook.xml.rels")).Root.Elements().Where(e => (string)e.Attribute("TargetMode") != "External").ToDictionary(e => (string)e.Attribute("Id"), e => (string)e.Attribute("Target"));
                var result = new List<ExcelSheet>(); int cellCount = 0;
                foreach (var sheet in Xml(zip.GetEntry("xl/workbook.xml")).Descendants(S + "sheet")) {
                    if ((string)sheet.Attribute("state") == "hidden" || (string)sheet.Attribute("state") == "veryHidden") continue;
                    string rel = (string)sheet.Attribute(XName.Get("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"));
                    if (!relationships.ContainsKey(rel)) continue;
                    var target = new Uri(new Uri("http://local/xl/workbook.xml"), relationships[rel]).AbsolutePath.TrimStart('/');
                    var entry = zip.GetEntry(Uri.UnescapeDataString(target)); if (entry == null) continue;
                    var xml = Xml(entry); var data = new ExcelSheet { Name = (string)sheet.Attribute("name") };
                    foreach (var row in xml.Descendants(S + "sheetData").Elements(S + "row")) foreach (var cell in row.Elements(S + "c")) {
                        if (++cellCount > 600000) throw new InvalidDataException("单元格过多，请按单位工程拆分后导入。");
                        var rc = Address((string)cell.Attribute("r")); string type = (string)cell.Attribute("t"), value = (string)cell.Element(S + "v") ?? "";
                        if (type == "s") { int index; value = int.TryParse(value, out index) && index >= 0 && index < strings.Count ? strings[index] : "#共享文本无效"; }
                        else if (type == "inlineStr") value = string.Concat(cell.Descendants(S + "t").Select(t => t.Value));
                        if (cell.Element(S + "f") != null && value == "") value = "#公式未缓存";
                        if (!data.Rows.ContainsKey(rc[0])) data.Rows[rc[0]] = new Dictionary<int, string>();
                        data.Rows[rc[0]][rc[1]] = value;
                    }
                    foreach (var merge in xml.Descendants(S + "mergeCell")) { var a = ((string)merge.Attribute("ref")).Split(':'); if (a.Length == 2) { var first = Address(a[0]); var last = Address(a[1]); data.Merges.Add(new[] { first[0], first[1], last[0], last[1] }); } }
                    result.Add(data);
                }
                return result;
            }
        }
    }
}
