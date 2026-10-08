using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace SettlementReview {
    public static partial class Core {
        public static void ParameterizeBasisSection(XDocument xml) {
            var body = xml.Root.Element(W + "body"); if (body == null) return;
            var paragraphs = body.Elements(W + "p").ToList();
            Func<XElement, string> heading = p => Regex.Replace(RawParagraphText(p).Trim(), @"^[一二三四五六七八九十\d\s、.．（）()]+", "").TrimEnd('：', ':');
            int start = paragraphs.FindIndex(p => heading(p) == "审核依据及原则");
            if (start < 0) return;
            int end = paragraphs.FindIndex(start + 1, p => heading(p) == "审核内容"); if (end < 0) return;
            var contents = paragraphs.Skip(start + 1).Take(end - start - 1).Where(p => RawParagraphText(p).Trim() != "").ToList();
            if (contents.Count == 0 || contents.Any(p => RawParagraphText(p).Contains("{{审核依据及原则}}"))) return;
            if (contents.Any(p => p.Descendants().Any(e => new[] { "drawing", "pict", "object", "sectPr", "fldSimple" }.Contains(e.Name.LocalName)))) return;
            SetParagraph(contents[0], "{{审核依据及原则}}");
            foreach (var p in contents.Skip(1)) p.Remove();
        }
        public static string TemplateBasis(byte[] source) {
            var paragraphs = TemplateParagraphs(source).Where(p => p.Part == "word/document.xml").Select(p => p.Text).ToList();
            int start = paragraphs.FindIndex(t => Regex.Replace(t, @"^[一二三四五六七八九十\d\s、.．（）()]+", "").TrimEnd('：', ':') == "审核依据及原则");
            if (start >= 0) {
                int end = paragraphs.FindIndex(start + 1, t => Regex.Replace(t, @"^[一二三四五六七八九十\d\s、.．（）()]+", "").TrimEnd('：', ':') == "审核内容");
                if (end > start) {
                    string text = string.Join("\r\n", paragraphs.Skip(start + 1).Take(end - start - 1).Where(t => !string.IsNullOrWhiteSpace(t)));
                    if (text != "" && !text.Contains("{{")) return text;
                }
            }
            return FixedBody["审核依据及原则"];
        }
        public static string ProjectBasis(Project project) {
            string text = Get(project, "审核依据及原则");
            if (text != "") return text;
            text = Get(project, "模板审核依据");
            return text == "" ? FixedBody["审核依据及原则"] : text;
        }
        public static bool IsAbsent(string value) { return new[] { "无", "不适用", "无代建单位" }.Contains((value ?? "").Trim()); }
        public static void ExpandBasisParagraphs(XDocument xml, Dictionary<string, string> values) {
            foreach (var p in xml.Descendants(W + "p").Where(p => RawParagraphText(p).Trim() == "{{审核依据及原则}}").ToList()) {
                if (p.Descendants().Any(e => new[] { "drawing", "pict", "object", "fldSimple" }.Contains(e.Name.LocalName))) continue;
                var replacements = new List<XElement>();
                foreach (string line in values["审核依据及原则"].Replace("\r\n", "\n").Split('\n')) {
                    var copy = new XElement(p); copy.Descendants(W + "sectPr").Remove();
                    string raw = RawParagraphText(copy); ReplaceSpan(copy, raw.IndexOf("{{审核依据及原则}}", StringComparison.Ordinal), "{{审核依据及原则}}".Length, line);
                    replacements.Add(copy);
                }
                var section = p.Descendants(W + "sectPr").FirstOrDefault();
                if (section != null) { var props = replacements.Last().Element(W + "pPr"); if (props == null) { props = new XElement(W + "pPr"); replacements.Last().AddFirst(props); } props.Add(new XElement(section)); }
                p.ReplaceWith(replacements);
            }
        }
        public static void RemoveAbsentProjectLines(XDocument xml, Project project) {
            var absent = new HashSet<string>(Basic.Where(f => IsAbsent(Get(project, f.Name))).Select(f => f.Name));
            foreach (var p in xml.Descendants(W + "p").ToList()) {
                string text = RawParagraphText(p);
                var keys = Regex.Matches(text, @"\{\{([^{}]+)\}\}").Cast<Match>().Select(m => m.Groups[1].Value).ToList();
                if (keys.Count == 0 || keys.Any(k => !absent.Contains(k))) continue;
                // Delete only an entire optional-field paragraph, never a paragraph with live fields.
                // Section properties must survive because they control following pages.
                var section = p.Element(W + "pPr") == null ? null : p.Element(W + "pPr").Element(W + "sectPr");
                if (section != null) { p.ReplaceNodes(new XElement(W + "pPr", new XElement(section))); continue; }
                if (p.Parent.Name == W + "tc" && p.Parent.Elements(W + "p").Count() == 1) {
                    var row = p.Ancestors(W + "tr").FirstOrDefault();
                    if (row != null && !row.Descendants(W + "vMerge").Any() && row.Elements(W + "tc").Count() <= 2 &&
                        row.Elements(W + "tc").All(c => { string v = string.Concat(c.Descendants(W + "t").Select(t => t.Value)).Trim(); return v == "" || v == text.Trim() || absent.Any(k => v.TrimEnd('：', ':') == k); })) row.Remove();
                    else p.RemoveNodes();
                } else p.Remove();
            }
        }
    }

    public static class ReviewGroupEditor {
        // A group name identifies the complete unit, including rows hidden by filters.
        public static void RenameGroup(ExcelReviewState state, IEnumerable<ReviewItem> selection, string name) {
            var rows = selection.ToList();
            var groups = new HashSet<string>(rows.Select(i => ExcelReview.Norm(i.Group)));
            var keys = new HashSet<string>(state.Items.Where(i => groups.Contains(ExcelReview.Norm(i.Group))).Select(i => i.CalculationKey).Where(k => !string.IsNullOrEmpty(k)));
            Rename(state, state.Items.Where(i => groups.Contains(ExcelReview.Norm(i.Group)) || !string.IsNullOrEmpty(i.CalculationKey) && keys.Contains(i.CalculationKey)).ToList(), name);
        }
        public static void RepairAssociations(ExcelReviewState state) {
            if (state.Totals == null || state.Balances == null) return;
            foreach (var total in state.Totals.Where(t => !string.IsNullOrEmpty(t.Key))) {
                var groups = state.Items.Where(i => i.CalculationKey == total.Key).Select(i => i.Group).Where(g => !string.IsNullOrWhiteSpace(g)).GroupBy(ExcelReview.Norm).Select(g => g.First()).ToList();
                if (groups.Count != 1 || ExcelReview.Norm(total.Group) == ExcelReview.Norm(groups[0])) continue;
                string old = total.Group; total.Group = groups[0];
                if (!state.Items.Any(i => ExcelReview.Norm(i.Group) == ExcelReview.Norm(old)))
                    foreach (var b in state.Balances.Where(b => ExcelReview.Norm(b.Group) == ExcelReview.Norm(old))) b.Group = groups[0];
            }
        }
        public static void Rename(ExcelReviewState state, IEnumerable<ReviewItem> selection, string name) {
            name = ExcelReview.CleanGroup(name);
            if (string.IsNullOrWhiteSpace(name)) throw new InvalidDataException("请填写单项工程名称。");
            var selected = new HashSet<ReviewItem>(selection.Where(i => state.Items.Contains(i)));
            if (selected.Count == 0) throw new InvalidDataException("请先选中需要改名的明细行。");
            ExcelReview.Normalize(state);
            var wholeGroups = selected.Select(i => i.Group).Distinct().Where(g => state.Items.Where(i => ExcelReview.Norm(i.Group) == ExcelReview.Norm(g)).All(selected.Contains)).ToList();
            wholeGroups.AddRange(state.Totals.Where(t => !string.IsNullOrEmpty(t.Key) && state.Items.Any(i => i.CalculationKey == t.Key) && state.Items.Where(i => i.CalculationKey == t.Key).All(selected.Contains) && !state.Items.Any(i => ExcelReview.Norm(i.Group) == ExcelReview.Norm(t.Group) && !selected.Contains(i))).Select(t => t.Group));
            var balances = state.Balances.Where(b => wholeGroups.Any(g => ExcelReview.Norm(g) == ExcelReview.Norm(b.Group)) || ExcelReview.Norm(b.Group) == ExcelReview.Norm(name)).ToList();
            if (balances.Count > 1 && balances.Any(b => b.NetOverride.HasValue)) throw new InvalidDataException("合并名称涉及人工总额，请先在金额配平中清除人工总额并核对，避免误合并。");
            if (balances.Where(b => b.Enabled).Select(b => b.Target).Distinct().Count() > 1) throw new InvalidDataException("合并名称涉及不同报告归属的补差，请先统一金额配平的报告归属。");
            foreach (var total in state.Totals) {
                var linked = state.Items.Where(i => i.CalculationKey == total.Key && !string.IsNullOrEmpty(total.Key)).ToList();
                if (wholeGroups.Any(g => ExcelReview.Norm(g) == ExcelReview.Norm(total.Group)) || linked.Count > 0 && linked.All(selected.Contains)) total.Group = name;
            }
            foreach (var b in state.Balances.Where(b => wholeGroups.Any(g => ExcelReview.Norm(g) == ExcelReview.Norm(b.Group)))) b.Group = name;
            foreach (var item in selected) item.Group = name;
            var duplicates = state.Balances.Where(b => ExcelReview.Norm(b.Group) == ExcelReview.Norm(name)).ToList();
            if (duplicates.Count > 1) {
                var keeper = duplicates.FirstOrDefault(b => b.Enabled) ?? duplicates[0];
                foreach (var b in duplicates.Where(b => !ReferenceEquals(b, keeper))) state.Balances.Remove(b);
            }
        }
    }
}
