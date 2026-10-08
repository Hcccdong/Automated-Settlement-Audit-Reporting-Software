using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace SettlementReview {
    public static partial class Core {
        static void SetWordValue(XElement parent, string name, string value) { var child = parent.Element(W + name); if (child == null) { child = new XElement(W + name); parent.Add(child); } child.SetAttributeValue(W + "val", value); }
        static int NextWordId(XElement root, string element, string attribute) { return root.Elements(W + element).Select(e => (int?)e.Attribute(W + attribute) ?? 0).DefaultIfEmpty(0).Max() + 1; }
        static XElement NumberingLevel(XElement donor, XDocument numbering) {
            var numPr = donor.Element(W + "pPr") == null ? null : donor.Element(W + "pPr").Element(W + "numPr");
            var numId = numPr == null ? null : numPr.Element(W + "numId");
            var num = numId == null ? null : numbering.Root.Elements(W + "num").FirstOrDefault(n => (string)n.Attribute(W + "numId") == (string)numId.Attribute(W + "val"));
            string abstractId = num == null || num.Element(W + "abstractNumId") == null ? null : (string)num.Element(W + "abstractNumId").Attribute(W + "val");
            var definition = numbering.Root.Elements(W + "abstractNum").FirstOrDefault(n => (string)n.Attribute(W + "abstractNumId") == abstractId);
            var level = definition == null ? null : definition.Elements(W + "lvl").FirstOrDefault();
            return level == null ? new XElement(W + "lvl", new XAttribute(W + "ilvl", "0"), new XElement(W + "start", new XAttribute(W + "val", "1")), new XElement(W + "numFmt", new XAttribute(W + "val", "decimal")), new XElement(W + "lvlText", new XAttribute(W + "val", "%1.")), new XElement(W + "pPr", new XElement(W + "tabs", new XElement(W + "tab", new XAttribute(W + "val", "num"), new XAttribute(W + "pos", "480"))), new XElement(W + "ind", new XAttribute(W + "left", "480"), new XAttribute(W + "hanging", "480")))) : new XElement(level);
        }
        static int AddReviewNumbering(XDocument numbering, XElement donor, string field) {
            int abstractId = NextWordId(numbering.Root, "abstractNum", "abstractNumId"), numId = NextWordId(numbering.Root, "num", "numId");
            var level = NumberingLevel(donor, numbering); level.SetAttributeValue(W + "ilvl", "0"); SetWordValue(level, "start", "1"); SetWordValue(level, "numFmt", "decimal"); SetWordValue(level, "lvlText", "%1.");
            var definition = new XElement(W + "abstractNum", new XAttribute(W + "abstractNumId", abstractId), new XElement(W + "multiLevelType", new XAttribute(W + "val", "singleLevel")), new XElement(W + "name", new XAttribute(W + "val", "SettlementReview_" + field)), level);
            var firstNum = numbering.Root.Element(W + "num"); if (firstNum == null) numbering.Root.Add(definition); else firstNum.AddBeforeSelf(definition);
            numbering.Root.Add(new XElement(W + "num", new XAttribute(W + "numId", numId), new XElement(W + "abstractNumId", new XAttribute(W + "val", abstractId)), new XElement(W + "lvlOverride", new XAttribute(W + "ilvl", "0"), new XElement(W + "startOverride", new XAttribute(W + "val", "1")))));
            return numId;
        }
        static XElement ReviewParagraph(XElement donor, string text, int? numId, bool heading, XElement listLevel) {
            var properties = donor.Element(W + "pPr") == null ? new XElement(W + "pPr") : new XElement(donor.Element(W + "pPr"));
            properties.Elements(W + "numPr").Remove(); properties.Elements(W + "sectPr").Remove(); properties.Elements(W + "outlineLvl").Remove();
            if (numId.HasValue) {
                var numbering = new XElement(W + "numPr", new XElement(W + "ilvl", new XAttribute(W + "val", "0")), new XElement(W + "numId", new XAttribute(W + "val", numId.Value)));
                var style = properties.Element(W + "pStyle"); if (style == null) properties.AddFirst(numbering); else style.AddAfterSelf(numbering);
                var levelInd = listLevel == null || listLevel.Element(W + "pPr") == null ? null : listLevel.Element(W + "pPr").Element(W + "ind");
                if (levelInd != null) { properties.Elements(W + "ind").Remove(); var after = properties.Element(W + "spacing"); if (after == null) properties.Add(new XElement(levelInd)); else after.AddAfterSelf(new XElement(levelInd)); }
            }
            if (heading) { properties.Elements(W + "keepNext").Remove(); var keep = new XElement(W + "keepNext"); var style = properties.Element(W + "pStyle"); if (style == null) properties.AddFirst(keep); else style.AddAfterSelf(keep); }
            var originalRun = donor.Elements(W + "r").FirstOrDefault(r => r.Descendants(W + "t").Any(t => t.Value.Trim() != ""));
            var runProperties = originalRun == null ? null : originalRun.Element(W + "rPr");
            if (runProperties == null) runProperties = properties.Element(W + "rPr");
            var run = new XElement(W + "r"); if (runProperties != null) run.Add(new XElement(runProperties));
            if (heading) { var rp = run.Element(W + "rPr"); if (rp == null) { rp = new XElement(W + "rPr"); run.AddFirst(rp); } SetWordValue(rp, "b", "1"); }
            string[] order = { "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange" };
            properties.ReplaceNodes(properties.Elements().OrderBy(e => Array.IndexOf(order, e.Name.LocalName) < 0 ? 99 : Array.IndexOf(order, e.Name.LocalName)).ToList());
            run.Add(new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text));
            return new XElement(W + "p", properties, run);
        }
        public static byte[] FormatReviewTemplate(Project project, byte[] bytes, Dictionary<string, string> values) {
            if (project.Excel == null || !project.Excel.AutoLink && string.IsNullOrWhiteSpace(project.Excel.LastText)) return bytes;
            var managed = (project.Excel.AutoLink ? ExcelReview.Targets : new[] { project.Excel.LastTarget }).Where(f => values.ContainsKey(f) && !IsAbsent(values[f])).ToArray();
            using (var input = new MemoryStream(bytes)) using (var zip = new ZipArchive(input, ZipArchiveMode.Read)) {
                var document = Xml(zip.GetEntry("word/document.xml")); var paragraphs = document.Descendants(W + "p").ToList();
                var matterDonor = paragraphs.FirstOrDefault(p => RawParagraphText(p).Contains("{{基期价格}}"));
                var numberEntry = zip.GetEntry("word/numbering.xml"); var numbering = numberEntry == null ? new XDocument(new XElement(W + "numbering")) : Xml(numberEntry);
                var numbers = new Dictionary<string, int>(); bool changed = false;
                foreach (var paragraph in paragraphs) {
                    string original = RawParagraphText(paragraph);
                    if (project.Excel.AutoLink) original = original.Replace("{{修正造价情况说明}}", string.Join("\n", ExcelReview.Targets.Where(f => !IsAbsent(values[f])).Select(f => f.Replace("说明", "") + "：\n{{" + f + "}}")));
                    if (!managed.Any(f => original.Contains("{{" + f + "}}"))) continue;
                    // Leave complex inline objects in place; only plain text placeholders become repeated paragraphs.
                    if (paragraph.Descendants().Any(e => new[] { "drawing", "pict", "object", "fldSimple", "hyperlink", "bookmarkStart", "bookmarkEnd" }.Contains(e.Name.LocalName))) continue;
                    var donor = matterDonor ?? paragraph; var level = NumberingLevel(donor, numbering); var replacement = new List<XElement>();
                    string pattern = "(\\{\\{(?:" + string.Join("|", managed.Select(Regex.Escape)) + ")\\}\\})";
                    foreach (string segment in Regex.Split(original, pattern)) {
                        if (segment == "") continue;
                        string field = segment.StartsWith("{{") && segment.EndsWith("}}") ? segment.Substring(2, segment.Length - 4) : null;
                        if (field != null && managed.Contains(field)) {
                            foreach (string line in values[field].Replace("\r\n", "\n").Split('\n').Where(t => !string.IsNullOrWhiteSpace(t))) {
                                var numbered = Regex.Match(line, @"^\d+\.\s+(?<text>.+)$");
                                if (numbered.Success) { if (!numbers.ContainsKey(field)) numbers[field] = AddReviewNumbering(numbering, donor, field); replacement.Add(ReviewParagraph(donor, numbered.Groups["text"].Value, numbers[field], false, level)); }
                                else replacement.Add(ReviewParagraph(donor, line.Trim(), null, line.TrimEnd().EndsWith("："), level));
                            }
                        } else foreach (string line in segment.Replace("\r\n", "\n").Split('\n').Where(t => !string.IsNullOrWhiteSpace(t))) replacement.Add(ReviewParagraph(paragraph, line, null, false, null));
                    }
                    if (replacement.Count > 0) { paragraph.ReplaceWith(replacement); changed = true; }
                }
                if (!changed) return bytes;
                var updates = new Dictionary<string, XDocument> { { "word/document.xml", document } };
                if (numbers.Count > 0) {
                    updates["word/numbering.xml"] = numbering;
                    XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
                    var contentTypes = Xml(zip.GetEntry("[Content_Types].xml"));
                    if (!contentTypes.Root.Elements(ct + "Override").Any(e => (string)e.Attribute("PartName") == "/word/numbering.xml")) { contentTypes.Root.Add(new XElement(ct + "Override", new XAttribute("PartName", "/word/numbering.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"))); updates["[Content_Types].xml"] = contentTypes; }
                    XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships"; var relEntry = zip.GetEntry("word/_rels/document.xml.rels"); var relations = relEntry == null ? new XDocument(new XElement(rel + "Relationships")) : Xml(relEntry);
                    if (!relations.Root.Elements().Any(e => ((string)e.Attribute("Type") ?? "").EndsWith("/numbering"))) { relations.Root.Add(new XElement(rel + "Relationship", new XAttribute("Id", "rIdReview" + Guid.NewGuid().ToString("N")), new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"), new XAttribute("Target", "numbering.xml"))); updates["word/_rels/document.xml.rels"] = relations; }
                }
                using (var output = new MemoryStream()) { using (var target = new ZipArchive(output, ZipArchiveMode.Create, true)) {
                    foreach (var entry in zip.Entries) if (!updates.ContainsKey(entry.FullName)) using (var source = entry.Open()) using (var dest = target.CreateEntry(entry.FullName).Open()) source.CopyTo(dest);
                    foreach (var pair in updates) using (var writer = new StreamWriter(target.CreateEntry(pair.Key).Open(), new UTF8Encoding(false))) pair.Value.Save(writer, SaveOptions.DisableFormatting);
                } return output.ToArray(); }
            }
        }
        public static HashSet<string> ReviewNumberIds(ZipArchive zip) {
            var entry = zip.GetEntry("word/numbering.xml"); if (entry == null) return new HashSet<string>(); var xml = Xml(entry);
            var definitions = new HashSet<string>(xml.Root.Elements(W + "abstractNum").Where(e => e.Element(W + "name") != null && ((string)e.Element(W + "name").Attribute(W + "val") ?? "").StartsWith("SettlementReview_")).Select(e => (string)e.Attribute(W + "abstractNumId")));
            return new HashSet<string>(xml.Root.Elements(W + "num").Where(e => e.Element(W + "abstractNumId") != null && definitions.Contains((string)e.Element(W + "abstractNumId").Attribute(W + "val"))).Select(e => (string)e.Attribute(W + "numId")));
        }
    }
}
