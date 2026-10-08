using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace SettlementReview {
    public static partial class Core {
        public static readonly string[] DateFields = { "计划开工", "计划竣工", "实际开工", "实际竣工" };
        public static readonly string[] DateInputFields = DateFields.Concat(new[] { "报告日期" }).ToArray();
        public static readonly string[] MustEnter = { "项目名称", "报告编号", "报告日期", "送审金额", "审定金额", "核减合计", "审增合计", "净核减金额" };
        const string DatePattern = @"\d{4}年\d{1,2}月\d{1,2}日";
        public static DateTime? ParseDate(string text) {
            DateTime result;
            return DateTime.TryParseExact(text, new[] { "yyyy-MM-dd", "yyyy年M月d日", "yyyy/M/d" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out result) ? (DateTime?)result : null;
        }
        public static int CalendarDays(DateTime start, DateTime end, bool inclusive) {
            if (end.Date < start.Date) throw new InvalidDataException("竣工日期不能早于开工日期。");
            return (end.Date - start.Date).Days + (inclusive ? 1 : 0);
        }
        static string SchedulePattern(string text) {
            string result = text;
            var dates = new Dictionary<string, string> { { "计划开工", "计划开工" }, { "计划竣工(?:验收)?", "计划竣工" }, { "实际开工", "实际开工" }, { "实际竣工(?:验收)?", "实际竣工" } };
            foreach (var pair in dates) result = Regex.Replace(result, "(" + pair.Key + @"(?:日期)?(?:为|：|:)?\s*)(" + DatePattern + ")", m => m.Groups[1].Value + "{{" + pair.Value + "}}");
            result = Regex.Replace(result, @"((?:合同总工期|合同工期)(?:为)?\s*)\d+", "$1{{合同工期日历天数}}");
            result = Regex.Replace(result, @"(实际工期(?:为)?\s*)\d+", "$1{{实际工期日历天数}}");
            result = Regex.Replace(result, @"(工期延期(?:为)?\s*)\d+", "$1{{工期延期天数}}");
            return result;
        }
        public static void MapScheduleParagraph(XElement paragraph) {
            string original = RawParagraphText(paragraph); string pattern = SchedulePattern(original);
            var slotMatches = Regex.Matches(pattern, @"\{\{([^{}]+)\}\}").Cast<Match>().ToList();
            if (slotMatches.Count == 0) { SetParagraph(paragraph, "{{工期情况}}"); return; }
            string matcher = "^"; int cursor = 0;
            for (int i = 0; i < slotMatches.Count; i++) { var m = slotMatches[i]; matcher += Regex.Escape(pattern.Substring(cursor, m.Index - cursor)) + "(?<v" + i + ">.*?)"; cursor = m.Index + m.Length; }
            matcher += Regex.Escape(pattern.Substring(cursor)) + "$";
            var values = Regex.Match(original, matcher, RegexOptions.Singleline);
            for (int i = slotMatches.Count - 1; i >= 0; i--) { var g = values.Groups["v" + i]; if (g.Success && g.Length > 0) ReplaceSpan(paragraph, g.Index, g.Length, slotMatches[i].Value); }
            string raw = RawParagraphText(paragraph); var tail = Regex.Match(raw, @"(?<=\{\{工期延期天数\}\}天[，,])(?<tail>.+)$");
            if (tail.Success) ReplaceSpan(paragraph, tail.Groups["tail"].Index, tail.Groups["tail"].Length, "{{工期补充说明}}");
        }
        public static Dictionary<string, string> TemplateDefaults(byte[] source, byte[] prepared) {
            var result = new Dictionary<string, string>();
            var oldParagraphs = TemplateParagraphs(source); var newParagraphs = TemplateParagraphs(prepared);
            // Match unchanged wording around fields; prefer matching persistent paragraph IDs where available.
            using (var sourceStream = new MemoryStream(source)) using (var a = new ZipArchive(sourceStream, ZipArchiveMode.Read))
            using (var preparedStream = new MemoryStream(prepared)) using (var b = new ZipArchive(preparedStream, ZipArchiveMode.Read)) {
                foreach (var entry in b.Entries.Where(e => IsTextPart(e.FullName))) {
                    var oldEntry = a.GetEntry(entry.FullName); if (oldEntry == null) continue;
                    var originals = Xml(oldEntry).Descendants(W + "p").ToList();
                    foreach (var p in Xml(entry).Descendants(W + "p")) {
                        string pattern = RawParagraphText(p); var marks = Regex.Matches(pattern, @"\{\{([^{}]+)\}\}").Cast<Match>().ToList(); if (marks.Count == 0) continue;
                        var id = p.Attributes().FirstOrDefault(x => x.Name.LocalName == "paraId");
                        IEnumerable<XElement> candidates = originals;
                        if (id != null) candidates = originals.Where(o => (string)o.Attribute(id.Name) == id.Value);
                        string regex = "^"; int pos = 0;
                        for (int i = 0; i < marks.Count; i++) { regex += Regex.Escape(pattern.Substring(pos, marks[i].Index - pos)) + "(?<g" + i + ">.*?)"; pos = marks[i].Index + marks[i].Length; }
                        regex += Regex.Escape(pattern.Substring(pos)) + "$";
                        var matches = candidates.Select(o => Regex.Match(RawParagraphText(o), regex, RegexOptions.Singleline)).Where(m => m.Success).ToList();
                        if (matches.Count != 1) continue;
                        for (int i = 0; i < marks.Count; i++) { string value = matches[0].Groups["g" + i].Value.Trim(); if (!value.Contains("{{") && value != "") result[marks[i].Groups[1].Value] = value; }
                    }
                }
            }
            string duration = oldParagraphs.Where(p => p.Part == "word/document.xml").Select(p => p.Text).FirstOrDefault(t => t.Contains("计划开工") && t.Contains("实际开工"));
            if (duration != null) {
                result["工期情况"] = duration; result["工期句式"] = SchedulePattern(duration);
                var tail = Regex.Match(result["工期句式"], @"(?<=\{\{工期延期天数\}\}天[，,])(?<tail>.+)$");
                if (tail.Success) { result["工期补充说明"] = tail.Groups["tail"].Value; result["工期句式"] = result["工期句式"].Substring(0, tail.Groups["tail"].Index) + "{{工期补充说明}}"; }
            }
            return result;
        }
        public static void ApplyTemplateDefaults(Project project, byte[] source, byte[] prepared, bool refresh = false) {
            string basis = TemplateBasis(source); project.Values["模板审核依据"] = basis;
            if (refresh || Get(project, "审核依据及原则") == "") project.Values["审核依据及原则"] = basis;
            var refreshFields = new HashSet<string>(Basic.Concat(Schedule).Select(f => f.Name).Concat(MatterFields).Concat(new[] { "工期情况", "工期句式" }));
            foreach (var pair in TemplateDefaults(source, prepared)) {
                if (MustEnter.Contains(pair.Key) && !(refresh && refreshFields.Contains(pair.Key))) continue;
                if (pair.Key == "工期句式") { project.Values[pair.Key] = pair.Value; continue; }
                if (Get(project, pair.Key) == "" || refresh && refreshFields.Contains(pair.Key)) project.Values[pair.Key] = pair.Value;
            }
            if (Get(project, "工期口径") == "") project.Values["工期口径"] = "首尾均计";
        }
        public static Dictionary<string, string> ScheduleValues(Project project, bool draft) {
            var values = new Dictionary<string, string>();
            foreach (string key in DateFields) {
                string text = Get(project, key); var date = ParseDate(text);
                if (text != "" && !date.HasValue) throw new InvalidDataException(key + "日期无效，请使用日期选择框。");
                values[key] = date.HasValue ? date.Value.ToString("yyyy年M月d日") : "待填写";
            }
            bool inclusive = Get(project, "工期口径") != "日期相减";
            foreach (string prefix in new[] { "计划", "实际" }) {
                var start = ParseDate(Get(project, prefix + "开工")); var end = ParseDate(Get(project, prefix + "竣工"));
                values[prefix + "工期日历天数"] = start.HasValue && end.HasValue ? CalendarDays(start.Value, end.Value, inclusive).ToString() : "待填写";
            }
            string contract = Get(project, "合同工期日历天数"); int contractDays;
            if (contract != "" && (!int.TryParse(contract, out contractDays) || contractDays < 0)) throw new InvalidDataException("合同工期请填写非负整数日历天数。");
            values["合同工期日历天数"] = contract == "" ? "待填写" : contract;
            int actual; values["工期延期天数"] = int.TryParse(values["实际工期日历天数"], out actual) && int.TryParse(contract, out contractDays) ? Math.Max(0, actual - contractDays).ToString() : "待填写";
            values["工期补充说明"] = Get(project, "工期补充说明");
            string pattern = Get(project, "工期句式");
            if (pattern == "") pattern = "本工程合同总工期{{合同工期日历天数}}日历天，计划开工日期为{{计划开工}}，计划竣工日期为{{计划竣工}}，实际开工日期为{{实际开工}}，实际竣工日期为{{实际竣工}}，实际工期为{{实际工期日历天数}}日历天。{{工期补充说明}}";
            values["工期情况"] = Regex.Replace(pattern, @"\{\{([^{}]+)\}\}", m => values.ContainsKey(m.Groups[1].Value) ? values[m.Groups[1].Value] : m.Value);
            return values;
        }
        public static decimal SignedMoney(string value) { return value.StartsWith("-") ? -Money(value.Substring(1), "净核减金额") : Money(value, "净核减金额"); }
        public static string AmountCheck(Project p) {
            string before = Get(p, "送审金额"), after = Get(p, "审定金额"), entered = Get(p, "净核减金额");
            if (before == "" || after == "") return "请先填写送审金额和审核金额。";
            decimal expected = Money(before, "送审金额") - Money(after, "审核金额");
            string expectedLabel = expected < 0 ? "净核增" : "净核减";
            var lines = new List<string> { "送审 − 审核：" + expectedLabel + " " + F(Math.Abs(expected)) + " 元" };
            if (entered == "") lines.Add("将自动生成净核减或净核增金额，也可以切换为手动填写。");
            else { decimal signed = SignedMoney(entered), difference = signed - expected; string enteredLabel = signed < 0 ? "净核增" : "净核减"; lines.Add(difference == 0 ? enteredLabel + "金额核对一致。" : "提示：所填" + enteredLabel + "金额与计算值相差 " + F(difference) + " 元。"); }
            string decrease = Get(p, "核减合计"), increase = Get(p, "审增合计");
            if (decrease != "" && increase != "") { decimal net = Money(decrease, "核减金额") - Money(increase, "核增金额"); lines.Add(net == expected ? "核减 − 核增核对一致。" : "提示：核减 − 核增 = " + F(net) + " 元，与送审 − 审核相差 " + F(net - expected) + " 元。"); }
            return string.Join("\n", lines);
        }
    }
}
