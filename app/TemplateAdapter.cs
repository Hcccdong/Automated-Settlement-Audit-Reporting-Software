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
        public static readonly Dictionary<string, string> FixedBody = new Dictionary<string, string> {
            { "审核依据及原则", "《浙江省建设工程计价规则》（2018版）\n《浙江省房屋建筑与装饰工程预算定额》（2018版）\n《浙江省房屋建筑安装工程修缮预算定额》（2018版）\n《浙江省通用安装工程预算定额》（2018版）\n《浙江省市政工程预算定额》（2018版）\n《浙江省园林绿化养护预算定额》（2018版）\n《工具式外墙脚手架浙江省补充定额》浙建站计〔2024〕4号\n《关于颁发浙江省建筑工程计价依据（2018版）的通知》浙建建〔2018〕61号\n《关于增值税调整后我省建设工程计价依据增值税税率及有关计价调整的通知》浙建建发〔2019〕92号\n工程相关纪要及其他相关政策性文件；\n送审单位提供的合同、招标文件、竣工图等其他相关资料。" },
            { "审核内容", "工程量、联系单、取费标准、材料价格等其他与造价相关的内容。" },
            { "审核程序及方法", "审核程序：明确审核范围,拟定结算实施方案,收集审核所需资料,初步确定结算审核结果,复核审核结果,撰写审核报告,审核报告审定、会签,审核资料整理、归档。" }
        };
        static byte[] builtinCache;
        public static byte[] SourceTemplate() {
            using (var s = typeof(Core).Assembly.GetManifestResourceStream("UserReport.docx")) {
                if (s == null) throw new FileNotFoundException("软件内置模板缺失，请使用完整发布版本。");
                using (var ms = new MemoryStream()) { s.CopyTo(ms); return ms.ToArray(); }
            }
        }
        public static byte[] BuiltIn() { if (builtinCache == null) builtinCache = PrepareTemplate(SourceTemplate()); return (byte[])builtinCache.Clone(); }
        static string ParagraphText(XElement p) { return string.Concat(p.Descendants(W + "t").Select(t => t.Value)).Trim(); }
        public static string RawParagraphText(XElement p) { return string.Concat(p.Descendants(W + "t").Where(t => t.Ancestors(W + "p").First() == p).Select(t => t.Value)); }
        public static void ReplaceSpan(XElement p, int start, int length, string replacement) {
            var nodes = p.Descendants(W + "t").Where(t => t.Ancestors(W + "p").First() == p).ToList();
            if (start < 0 || length <= 0 || start + length > nodes.Sum(t => t.Value.Length)) throw new InvalidDataException("请选择要填写的文字。");
            int end = start + length, cursor = 0;
            foreach (var n in nodes) {
                int next = cursor + n.Value.Length;
                if (cursor < end && next > start) {
                    int a = Math.Max(0, start - cursor), b = Math.Min(n.Value.Length, end - cursor);
                    n.Value = n.Value.Substring(0, a) + (cursor <= start ? replacement : "") + n.Value.Substring(b);
                    n.SetAttributeValue(XNamespace.Xml + "space", "preserve");
                }
                cursor = next;
            }
        }
        static void ReplaceLiteral(XElement p, string oldText, string replacement) {
            if (string.IsNullOrEmpty(oldText)) return;
            string text = RawParagraphText(p);
            foreach (Match m in Regex.Matches(text, Regex.Escape(oldText)).Cast<Match>().Reverse()) ReplaceSpan(p, m.Index, m.Length, replacement);
        }
        static void SetParagraph(XElement p, string text) {
            // Edit text nodes only. Preserve drawings, borders, tabs, fields, bookmarks and every run property.
            string original = RawParagraphText(p);
            string leading = Regex.Match(original, @"^[ \t\u3000]*").Value;
            string trailing = Regex.Match(original, @"[ \t\u3000]*$").Value;
            int length = original.Length - leading.Length - trailing.Length;
            if (length > 0) ReplaceSpan(p, leading.Length, length, text);
            else p.Add(new XElement(W + "r", new XElement(W + "t", text)));
        }
        static void CleanRevisions(XDocument xml) {
            // Accept the current version, never the old formatting stored in change records.
            string[] changes = { "rPrChange", "pPrChange", "sectPrChange", "tblPrChange", "tblPrExChange", "trPrChange", "tcPrChange", "tblGridChange", "numberingChange" };
            xml.Descendants().Where(e => e.Name.Namespace == W && changes.Contains(e.Name.LocalName)).ToList().ForEach(e => e.Remove());
            foreach (var row in xml.Descendants(W + "tr").Where(r => r.Element(W + "trPr") != null && r.Element(W + "trPr").Element(W + "del") != null).ToList()) row.Remove();
            foreach (var cell in xml.Descendants(W + "tc").Where(c => c.Element(W + "tcPr") != null && c.Element(W + "tcPr").Element(W + "cellDel") != null).ToList()) cell.Remove();
            foreach (var p in xml.Descendants(W + "p").ToList()) {
                var pr = p.Element(W + "pPr"); var rp = pr == null ? null : pr.Element(W + "rPr");
                if (rp == null || rp.Element(W + "del") == null) continue;
                var next = p.ElementsAfterSelf().FirstOrDefault();
                if (next != null && next.Name == W + "p") { p.Add(next.Elements().Where(e => e.Name != W + "pPr").Select(e => new XElement(e))); next.Remove(); }
            }
            foreach (var e in xml.Descendants().Where(e => e.Name == W + "del" || e.Name == W + "moveFrom" || e.Name == W + "delText").ToList()) if (e.Parent != null) e.Remove();
            foreach (var e in xml.Descendants().Where(e => e.Name == W + "ins" || e.Name == W + "moveTo").Reverse().ToList()) if (e.Parent != null) e.ReplaceWith(e.Nodes().ToList());
            string[] remove = { "documentProtection", "writeProtection", "trackRevisions", "lock", "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd", "cellIns", "cellDel", "customXmlInsRangeStart", "customXmlInsRangeEnd", "customXmlDelRangeStart", "customXmlDelRangeEnd" };
            xml.Descendants().Where(e => e.Name.Namespace == W && remove.Contains(e.Name.LocalName)).Remove();
        }
        static bool MapReport(XDocument xml) {
            var body = xml.Root.Element(W + "body"); if (body == null) return false;
            var paras = body.Elements(W + "p").ToList();
            string[] headings = { "项目概况", "审核依据及原则", "审核内容", "审核程序及方法", "有关事项说明", "修正造价情况说明", "审核结论" };
            var anchors = new List<XElement>(); int last = -1;
            foreach (string heading in headings) {
                int index = paras.FindIndex(last + 1, p => Regex.Replace(ParagraphText(p), @"^[一二三四五六七八九十\d\s、.．（）()]+", "").TrimEnd('：', ':') == heading);
                if (index < 0) return false; anchors.Add(paras[index]); last = index;
            }
            Func<XElement, XElement, List<XElement>> between = (a, b) => paras.Skip(paras.IndexOf(a) + 1).Take(paras.IndexOf(b) - paras.IndexOf(a) - 1).Where(p => ParagraphText(p) != "").ToList();
            // Replace each dynamic section using its existing first paragraph; fixed sections stay untouched.
            Action<List<XElement>, string> replaceGroup = (group, text) => {
                if (group.Count == 0) throw new InvalidDataException("模板的可变内容段落缺失，无法定位填写位置。");
                SetParagraph(group[0], text); foreach (var p in group.Skip(1)) p.Remove();
            };
            var front = paras.Take(paras.IndexOf(anchors[0])).Where(p => ParagraphText(p) != "").ToList();
            if (front.Count < 6) return false;
            var title = front.FirstOrDefault(p => ParagraphText(p).StartsWith("关于") && ParagraphText(p).EndsWith("的"));
            var salutation = front.FirstOrDefault(p => ParagraphText(p).EndsWith("：") && !ParagraphText(p).StartsWith("我公司"));
            var intro = front.FirstOrDefault(p => ParagraphText(p).StartsWith("我公司接受"));
            if (title == null || salutation == null || intro == null) return false;
            string oldProject = ParagraphText(title).Substring(2); oldProject = oldProject.Substring(0, oldProject.Length - 1);
            string oldClient = ParagraphText(salutation).TrimEnd('：', ':');
            SetParagraph(front[0], "{{审核单位}}");
            ReplaceLiteral(front[1], oldProject, "{{项目名称}}");
            string oldNumber = ParagraphText(front[1]); int divider = oldNumber.IndexOf("{{项目名称}}", StringComparison.Ordinal);
            if (divider > 0) { string prefix = oldNumber.Substring(0, divider).TrimEnd('-', '－', '—', ' '); ReplaceLiteral(front[1], prefix, "{{报告编号}}"); }
            ReplaceLiteral(title, oldProject, "{{项目名称}}"); ReplaceLiteral(salutation, oldClient, "{{委托单位}}");
            ReplaceLiteral(intro, oldProject, "{{项目名称}}"); ReplaceLiteral(intro, oldClient, "{{委托单位}}");
            var basisParagraphs = between(anchors[1], anchors[2]);
            if (basisParagraphs.Count > 0) replaceGroup(basisParagraphs, "{{审核依据及原则}}");
            var basics = between(anchors[0], anchors[1]);
            foreach (string key in new[] { "批复总投资", "建设单位", "代建单位", "施工单位", "监理单位" }) {
                var p = basics.FirstOrDefault(x => ParagraphText(x).StartsWith(key + "：") || ParagraphText(x).StartsWith(key + ":"));
                if (p == null) return false;
                string raw = RawParagraphText(p); int begin = Math.Max(raw.IndexOf('：'), raw.IndexOf(':')) + 1;
                while (begin < raw.Length && char.IsWhiteSpace(raw[begin])) begin++;
                int end = raw.TrimEnd('；', ';', ' ', '\t', '\u3000').Length;
                if (end <= begin) return false; ReplaceSpan(p, begin, end - begin, "{{" + key + "}}");
            }
            var contract = basics.FirstOrDefault(x => ParagraphText(x).StartsWith("本工程合同额"));
            var period = basics.FirstOrDefault(x => ParagraphText(x).StartsWith("本工程合同总工期"));
            if (contract == null || period == null) return false;
            string contractText = RawParagraphText(contract);
            var content = Regex.Match(contractText, @"工程内容[:：](?<body>.+)");
            if (!content.Success) return false;
            ReplaceSpan(contract, content.Groups["body"].Index, content.Groups["body"].Length, "{{工程概况}}");
            var contractAmount = Regex.Match(RawParagraphText(contract), @"本工程合同额(?<amount>[\d,]+(?:\.\d+)?)元");
            if (!contractAmount.Success) return false;
            ReplaceSpan(contract, contractAmount.Groups["amount"].Index, contractAmount.Groups["amount"].Length, "{{合同金额}}"); MapScheduleParagraph(period);
            var matters = between(anchors[4], anchors[5]);
            if (matters.Count == 5) {
                for (int i = 0; i < 5; i++) {
                    var rate = Regex.Match(RawParagraphText(matters[i]), @"\d+(?:\.\d+)?[%％]");
                    if (i == 1 && rate.Success) ReplaceSpan(matters[i], rate.Index, rate.Length, "{{工程下浮率}}");
                    else SetParagraph(matters[i], "{{" + MatterFields[i] + "}}");
                }
            } else replaceGroup(matters, "{{有关事项说明}}");
            var adjustments = between(anchors[5], anchors[6]);
            string[] labels = { "合同标内", "合同标外", "联系单工程量" };
            var adjustmentHeadings = labels.Select(label => adjustments.FirstOrDefault(p => ParagraphText(p).StartsWith(label))).ToList();
            if (adjustmentHeadings.Any(p => p == null)) return false;
            MapMoneySentence(adjustments[0]);
            string[] fields = { "合同标内调整说明", "合同标外调整说明", "联系单调整说明" };
            for (int i = 0; i < fields.Length; i++) {
                var start = adjustmentHeadings[i]; var end = i < 2 ? adjustmentHeadings[i + 1] : anchors[6];
                replaceGroup(between(start, end), "{{" + fields[i] + "}}" + (i == 2 ? "\n其他调整：{{其他调整说明}}" : ""));
            }
            var after = paras.Skip(paras.IndexOf(anchors[6]) + 1).Where(p => ParagraphText(p) != "").ToList();
            if (after.Count < 6) return false;
            var attachment = after.FirstOrDefault(p => ParagraphText(p).StartsWith("附件"));
            var date = after.FirstOrDefault(p => Regex.IsMatch(ParagraphText(p), @"^\d{4}年\d{1,2}月\d{1,2}日$"));
            var cc = after.FirstOrDefault(p => ParagraphText(p).StartsWith("抄送"));
            if (attachment == null || date == null || cc == null) return false;
            foreach (var p in after.Take(after.IndexOf(attachment))) MapMoneySentence(p);
            int dateIndex = after.IndexOf(date); if (dateIndex < 1) return false;
            var signature = after[dateIndex - 1];
            replaceGroup(after.Skip(after.IndexOf(attachment)).Take(after.IndexOf(signature) - after.IndexOf(attachment)).ToList(), "附件：{{附件清单}}");
            SetParagraph(signature, "{{审核单位}}"); SetParagraph(date, "{{报告日期}}"); SetParagraph(cc, "抄送：{{抄送单位}}");
            return true;
        }
        static void MapMoneySentence(XElement p) {
            string text = RawParagraphText(p);
            string pattern = @"(?<label>净核减(?:[（(]增[）)])?(?:金额)?|建设单位审定金额|建设单位审核价|送审金额|审定金额|审核金额|核增金额|核减金额|核增|核减)(?:为|：|:)?\s*(?<amount>[\d,]+(?:\.\d+)?)\s*元";
            foreach (Match m in Regex.Matches(text, pattern).Cast<Match>().Reverse()) {
                string label = m.Groups["label"].Value;
                string field = label.StartsWith("净") ? "净核减金额" : label.StartsWith("核增") ? "核增金额" : label.StartsWith("核减") ? "核减金额" : label.StartsWith("建设") || label.StartsWith("送审") ? "送审金额" : "审核金额";
                var amount = m.Groups["amount"]; ReplaceSpan(p, amount.Index, amount.Length, "{{" + field + "}}");
            }
        }
        static void MapNetDirectionLabels(XDocument xml) {
            foreach (var p in xml.Descendants(W + "p")) {
                string text = RawParagraphText(p);
                var matches = Regex.Matches(text, @"净核减(?:[（(]增[）)])?(?:金额)?(?=\s*(?:为|：|:)?\s*\{\{净核减金额\}\})").Cast<Match>().ToList();
                foreach (var m in matches.AsEnumerable().Reverse()) ReplaceSpan(p, m.Index, m.Length, "{{净额方向}}" + (m.Value.Contains("金额") ? "金额" : ""));
            }
        }
        public static byte[] PrepareTemplate(byte[] bytes) {
            using (var input = new MemoryStream(bytes)) using (var zip = new ZipArchive(input, ZipArchiveMode.Read)) using (var output = new MemoryStream()) {
                if (zip.GetEntry("word/document.xml") == null) throw new InvalidDataException("请选择另存为 .docx 的Word报告。");
                if (zip.Entries.Sum(e => e.Length) > 32000000 || zip.Entries.Count > 3000) throw new InvalidDataException("模板过大，请使用32MB以内的模板。");
                using (var result = new ZipArchive(output, ZipArchiveMode.Create, true)) {
                    foreach (var entry in zip.Entries) {
                        var target = result.CreateEntry(entry.FullName);
                        using (var stream = target.Open()) {
                            if (!IsTextPart(entry.FullName)) { using (var source = entry.Open()) source.CopyTo(stream); continue; }
                            var xml = Xml(entry); string before = xml.ToString(SaveOptions.DisableFormatting); CleanRevisions(xml);
                            if (entry.FullName == "word/settings.xml") {
                                var update = xml.Root.Element(W + "updateFields");
                                if (update == null) { update = new XElement(W + "updateFields"); xml.Root.Add(update); }
                                update.SetAttributeValue(W + "val", "true");
                            }
                            if (entry.FullName == "word/document.xml" && !xml.Descendants(W + "p").Any(p => ParagraphText(p).Contains("{{"))) {
                                var candidate = new XDocument(xml);
                                if (MapReport(candidate)) xml = candidate;
                            }
                            if (entry.FullName == "word/document.xml") { ParameterizeBasisSection(xml); MapNetDirectionLabels(xml); }
                            if (before == xml.ToString(SaveOptions.DisableFormatting)) { using (var source = entry.Open()) source.CopyTo(stream); }
                            else using (var writer = new StreamWriter(stream, new UTF8Encoding(false))) xml.Save(writer, SaveOptions.DisableFormatting);
                        }
                    }
                } return output.ToArray();
            }
        }
    }
}
