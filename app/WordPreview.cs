using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;

namespace SettlementReview {
    public static class WordPreviewRenderer {
        static XDocument Read(ZipArchive zip, string name) {
            var entry = zip.GetEntry(name); if (entry == null) return null;
            using (var stream = entry.Open()) return XDocument.Load(stream);
        }
        static string Val(XElement element, string attribute = "val") { return element == null ? null : (string)element.Attribute(Core.W + attribute); }
        static XElement Property(XElement direct, Dictionary<string, XElement> styles, string styleId, string group, string name) {
            var own = direct == null ? null : direct.Element(Core.W + name); if (own != null) return own;
            var seen = new HashSet<string>();
            while (!string.IsNullOrEmpty(styleId) && styles.ContainsKey(styleId) && seen.Add(styleId)) {
                var style = styles[styleId]; var props = style.Element(Core.W + group); var found = props == null ? null : props.Element(Core.W + name); if (found != null) return found;
                styleId = Val(style.Element(Core.W + "basedOn"));
            }
            return null;
        }
        static bool Toggle(XElement value) { string v = Val(value); return value != null && v != "0" && v != "false" && v != "off" && v != "none"; }
        static int Twips(XElement element, string attribute) { int value; return int.TryParse(element == null ? null : (string)element.Attribute(Core.W + attribute), out value) ? value : 0; }
        static int Pixels(int twips) { return (int)Math.Round(twips / 15.0); }
        static Color TextColor(XElement color) { string value = Val(color); int rgb; return !string.IsNullOrEmpty(value) && value.Length == 6 && int.TryParse(value, System.Globalization.NumberStyles.HexNumber, null, out rgb) ? Color.FromArgb((rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255) : Color.Black; }
        static Font RunFont(XElement run, XElement paragraphProperties, Dictionary<string, XElement> styles, string paragraphStyle, XElement defaults, Font fallback) {
            var direct = run.Element(Core.W + "rPr"); string runStyle = Val(direct == null ? null : direct.Element(Core.W + "rStyle"));
            Func<string, XElement> prop = name => Property(direct, styles, runStyle, "rPr", name) ?? Property(paragraphProperties == null ? null : paragraphProperties.Element(Core.W + "rPr"), styles, paragraphStyle, "rPr", name) ?? (defaults == null ? null : defaults.Element(Core.W + name));
            var fonts = prop("rFonts"); string family = fonts == null ? fallback.FontFamily.Name : (string)fonts.Attribute(Core.W + "eastAsia") ?? (string)fonts.Attribute(Core.W + "ascii") ?? (string)fonts.Attribute(Core.W + "hAnsi") ?? fallback.FontFamily.Name;
            int halfPoints; float size = int.TryParse(Val(prop("sz")), out halfPoints) ? Math.Max(7F, Math.Min(48F, halfPoints / 2F)) : fallback.Size;
            FontStyle style = FontStyle.Regular; if (Toggle(prop("b"))) style |= FontStyle.Bold; if (Toggle(prop("i"))) style |= FontStyle.Italic; if (Toggle(prop("u"))) style |= FontStyle.Underline; if (Toggle(prop("strike"))) style |= FontStyle.Strikeout;
            try { return new Font(family, size, style, GraphicsUnit.Point); } catch { return new Font(fallback.FontFamily, size, style, GraphicsUnit.Point); }
        }
        static string NumberPrefix(XElement paragraph, XDocument numbering, Dictionary<string, int> counters) {
            var numPr = paragraph.Element(Core.W + "pPr") == null ? null : paragraph.Element(Core.W + "pPr").Element(Core.W + "numPr");
            string numId = Val(numPr == null ? null : numPr.Element(Core.W + "numId")); if (string.IsNullOrEmpty(numId)) return "";
            int count; counters[numId] = counters.TryGetValue(numId, out count) ? count + 1 : 1;
            if (numbering == null) return counters[numId] + ". ";
            var num = numbering.Root.Elements(Core.W + "num").FirstOrDefault(n => (string)n.Attribute(Core.W + "numId") == numId);
            string abstractId = Val(num == null ? null : num.Element(Core.W + "abstractNumId"));
            var definition = numbering.Root.Elements(Core.W + "abstractNum").FirstOrDefault(n => (string)n.Attribute(Core.W + "abstractNumId") == abstractId);
            var level = definition == null ? null : definition.Elements(Core.W + "lvl").FirstOrDefault(); string format = Val(level == null ? null : level.Element(Core.W + "lvlText"));
            return string.IsNullOrEmpty(format) ? counters[numId] + ". " : format.Replace("%1", counters[numId].ToString()) + " ";
        }
        public static void Render(RichTextBox box, byte[] docx) {
            box.Clear(); box.BackColor = Color.White; box.ForeColor = Color.Black;
            using (var memory = new MemoryStream(docx)) using (var zip = new ZipArchive(memory, ZipArchiveMode.Read)) {
                var document = Read(zip, "word/document.xml"); if (document == null) throw new InvalidDataException("Word文档缺少正文内容。");
                var styleDoc = Read(zip, "word/styles.xml"); var numbering = Read(zip, "word/numbering.xml");
                var styles = styleDoc == null ? new Dictionary<string, XElement>() : styleDoc.Root.Elements(Core.W + "style").Where(s => s.Attribute(Core.W + "styleId") != null).ToDictionary(s => (string)s.Attribute(Core.W + "styleId"), s => s);
                var defaults = styleDoc == null ? null : styleDoc.Root.Element(Core.W + "docDefaults"); defaults = defaults == null || defaults.Element(Core.W + "rPrDefault") == null ? null : defaults.Element(Core.W + "rPrDefault").Element(Core.W + "rPr");
                var counters = new Dictionary<string, int>(); var fallback = new Font("SimSun", 12F, FontStyle.Regular, GraphicsUnit.Point);
                foreach (var paragraph in document.Descendants(Core.W + "body").Elements(Core.W + "p")) {
                    string paragraphText = string.Concat(paragraph.Descendants(Core.W + "t").Select(t => t.Value)).Replace(" ", "").Replace("　", "");
                    if (paragraphText == "草稿待审核确认") continue;
                    var pPr = paragraph.Element(Core.W + "pPr"); string styleId = Val(pPr == null ? null : pPr.Element(Core.W + "pStyle"));
                    string align = Val(Property(pPr, styles, styleId, "pPr", "jc"));
                    box.SelectionAlignment = align == "center" ? HorizontalAlignment.Center : align == "right" || align == "end" ? HorizontalAlignment.Right : HorizontalAlignment.Left;
                    var indent = Property(pPr, styles, styleId, "pPr", "ind"); box.SelectionIndent = Math.Max(0, Pixels(Twips(indent, "left")) + 24); box.SelectionRightIndent = Math.Max(0, Pixels(Twips(indent, "right")) + 24);
                    int first = Twips(indent, "firstLine"), hanging = Twips(indent, "hanging"); box.SelectionHangingIndent = hanging > 0 ? Pixels(hanging) : first < 0 ? Pixels(-first) : 0;
                    if (first > 0) box.SelectionIndent += Pixels(first);
                    string prefix = NumberPrefix(paragraph, numbering, counters); if (prefix != "") { box.SelectionFont = fallback; box.SelectionColor = Color.Black; box.AppendText(prefix); }
                    var runs = paragraph.Descendants(Core.W + "r").Where(r => r.Ancestors(Core.W + "p").FirstOrDefault() == paragraph).ToList();
                    foreach (var run in runs) {
                        using (var font = RunFont(run, pPr, styles, styleId, defaults, fallback)) box.SelectionFont = font;
                        var direct = run.Element(Core.W + "rPr"); string runStyle = Val(direct == null ? null : direct.Element(Core.W + "rStyle"));
                        box.SelectionColor = TextColor(Property(direct, styles, runStyle, "rPr", "color") ?? Property(pPr == null ? null : pPr.Element(Core.W + "rPr"), styles, styleId, "rPr", "color"));
                        foreach (var node in run.Elements()) {
                            if (node.Name == Core.W + "t" || node.Name == Core.W + "delText" || node.Name == Core.W + "instrText") box.AppendText(node.Value);
                            else if (node.Name == Core.W + "tab") box.AppendText("\t"); else if (node.Name == Core.W + "br" || node.Name == Core.W + "cr") box.AppendText("\n");
                        }
                    }
                    box.AppendText("\n");
                }
                fallback.Dispose();
            }
            box.Select(0, 0); box.ScrollToCaret();
        }
    }
}
