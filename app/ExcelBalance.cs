using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.Serialization;
using System.Text.RegularExpressions;

namespace SettlementReview {
    [DataContract] public class ReviewTotal {
        [DataMember] public string Group, Sheet, Key, Source, Message;
        [DataMember] public decimal? Sent, Audit, Delta;
        [DataMember] public List<string> DetailSheets = new List<string>();
    }
    [DataContract] public class ReviewBalance {
        [DataMember] public string Group, Target, Reason;
        [DataMember] public bool Enabled;
        [DataMember] public decimal? NetOverride;
    }
    public class BalanceResult {
        public decimal? TotalDelta;
        public decimal Listed, Omitted, NonDetail, Residual;
        public string Source, Issue;
    }
    public static partial class ExcelReview {
        public const string FeeReason = "由于取费基数的变化，相应费率对应金额发生变化，审核予以调整";
        static string SheetKey(string name) { var m = Regex.Match(name ?? "", @"^(\d+[_-]\d+)[_-]"); return m.Success ? m.Groups[1].Value : ""; }
        public static bool IsCalculation(ExcelSheet sheet) { return Norm(sheet.Name).Contains("计算表") || sheet.Rows.Take(3).Any(r => r.Value.Values.Any(v => Norm(v).Contains("结算审核计算表"))); }
        static decimal? TotalNumber(string text) {
            decimal value; return decimal.TryParse((text ?? "").Replace(",", "").Replace("，", "").Trim(), NumberStyles.Float | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out value) ? (decimal?)Round(value) : null;
        }
        public static List<ReviewTotal> ReadTotals(List<ExcelSheet> sheets) {
            var result = new List<ReviewTotal>();
            foreach (var sheet in sheets.Where(IsCalculation)) {
                string group = sheet.Rows.Keys.Select(r => ProjectTitle(sheet, r)).FirstOrDefault(g => g != "") ?? "未填写工程名称";
                var cols = new Dictionary<string, int>();
                foreach (var row in sheet.Rows.Where(r => r.Key <= 15)) {
                    var candidate = new Dictionary<string, int>();
                    foreach (var cell in row.Value) {
                        string text = Norm(cell.Value); string key = new[] { "送审", "送审金额", "送审造价" }.Contains(text) ? "s" : new[] { "审核", "审查", "审定", "审核金额", "审查金额", "审定金额" }.Contains(text) ? "a" : new[] { "核增", "核增金额" }.Contains(text) ? "i" : new[] { "核减", "核减金额" }.Contains(text) ? "d" : new[] { "净核减", "净核减金额" }.Contains(text) ? "n" : null;
                        if (key != null) candidate[key] = cell.Key;
                    }
                    if (candidate.Count > cols.Count) cols = candidate;
                }
                var candidates = sheet.Rows.Where(r => r.Value.Any(c => c.Key <= 3 && new[] { "建设工程造价", "工程总造价", "工程造价合计", "单项工程合计", "总造价" }.Contains(Norm(c.Value)))).ToList();
                var total = new ReviewTotal { Group = group, Sheet = sheet.Name, Key = sheet.Name, Source = sheet.Name, Message = "" };
                if (candidates.Count != 1) total.Message = candidates.Count == 0 ? "未找到唯一的建设工程造价总额行" : "存在多个总额行，请核对口径";
                else {
                    int row = candidates[0].Key;
                    Func<string, decimal?> number = k => cols.ContainsKey(k) ? TotalNumber(sheet.Cell(row, cols[k])) : null;
                    total.Sent = number("s"); total.Audit = number("a"); decimal? inc = number("i"), dec = number("d"), net = number("n");
                    decimal? netColumns = inc.HasValue || dec.HasValue ? (decimal?)((inc ?? 0) - (dec ?? 0)) : net.HasValue ? -net : null;
                    total.Delta = total.Sent.HasValue && total.Audit.HasValue ? Round(total.Audit.Value - total.Sent.Value) : netColumns;
                    if (!total.Delta.HasValue) total.Message = "总额缺少有效金额或公式缓存结果";
                    if (total.Delta.HasValue && netColumns.HasValue && total.Delta.Value != netColumns.Value) total.Message = "总额送审减审核与核增减栏目不一致";
                    if (net.HasValue && total.Delta.HasValue && -net.Value != total.Delta.Value) total.Message = "净核减栏目与总额计算不一致";
                    if (inc < 0 || dec < 0) total.Message = "核增、核减栏目不是非负金额";
                    if (cols.Values.Any(c => { string raw = sheet.Cell(row, c).Trim(); return raw != "" && raw != "-" && raw != "—" && !TotalNumber(raw).HasValue; })) total.Message = "总额栏存在无效数值或未缓存公式，请核对后重新保存Excel";
                    total.Source += " 第" + row + "行（" + string.Join("、", cols.OrderBy(c => c.Value).Select(c => ColumnAddress(c.Value) + row)) + "）";
                }
                if (sheet.Rows.Where(r => r.Key <= 4).Any(r => r.Value.Values.Any(v => Norm(v).Contains("万元")))) { total.Delta = null; total.Message = "计算表标注万元，请换算为元后导入"; }
                result.Add(total);
            }
            return result;
        }
        static string ColumnAddress(int column) { string name = ""; while (column > 0) { column--; name = (char)('A' + column % 26) + name; column /= 26; } return name; }
        static void LinkTotals(ExcelReviewState state, List<ExcelSheet> sheets) {
            ReviewTotal prior = null;
            foreach (var sheet in sheets) {
                if (IsCalculation(sheet)) { prior = state.Totals.FirstOrDefault(t => t.Sheet == sheet.Name); continue; }
                if (!sheet.Rows.Keys.Where(r => r <= 40).Any(r => { var c = Columns(sheet, r); return c.ContainsKey("name") && (c.ContainsKey("sa") || c.ContainsKey("aa") || c.ContainsKey("inc") || c.ContainsKey("dec")); })) continue;
                string group = sheet.Rows.Keys.Select(r => ProjectTitle(sheet, r)).FirstOrDefault(g => g != "") ?? "";
                string key = SheetKey(sheet.Name);
                var matches = key == "" ? new List<ReviewTotal>() : state.Totals.Where(t => SheetKey(t.Sheet) == key && Norm(t.Group) == Norm(group)).ToList();
                ReviewTotal match = matches.Count == 1 ? matches[0] : key == "" && prior != null && Norm(prior.Group) == Norm(group) ? prior : null;
                if (match == null) continue;
                match.DetailSheets.Add(sheet.Name);
                foreach (var item in state.Items.Where(i => i.Sheet == sheet.Name)) item.CalculationKey = match.Key;
            }
            foreach (var group in state.Items.Select(i => i.Group).Concat(state.Totals.Select(t => t.Group)).GroupBy(Norm)) {
                string name = group.First(); var totals = state.Totals.Where(t => Norm(t.Group) == group.Key).ToList();
                var targets = state.Items.Where(i => Norm(i.Group) == group.Key).Select(i => i.Target).Distinct().ToList();
                state.Balances.Add(new ReviewBalance { Group = name, Target = targets.Count == 1 ? targets[0] : Targets[0], Reason = FeeReason, Enabled = totals.Count > 0 });
            }
            foreach (var total in state.Totals) state.Messages.Add("总额来源：" + total.Source + "，单项工程：" + total.Group + "，净核减：" + Money(total.Delta.HasValue ? -total.Delta : null) + "元" + (total.Message == "" ? "" : "；" + total.Message));
        }
        public static BalanceResult CalculateBalance(ExcelReviewState state, ReviewBalance balance) {
            var totals = state.Totals.Where(t => Norm(t.Group) == Norm(balance.Group)).ToList();
            var items = state.Items.Where(i => Norm(i.Group) == Norm(balance.Group)).ToList();
            var listed = Selected(state).Where(i => Norm(i.Group) == Norm(balance.Group)).ToList();
            var result = new BalanceResult { Source = string.Join("\r\n", totals.Select(t => t.Source + "，净核减 " + Money(state, t.Delta.HasValue ? -t.Delta : null) + " 元")), Issue = "", Listed = listed.Sum(i => DisplayRound(state, i.Delta.Value)) };
            result.Omitted = items.Where(i => i.Delta.HasValue && !listed.Contains(i)).Sum(i => DisplayRound(state, i.Delta.Value));
            if (balance.NetOverride.HasValue) { result.TotalDelta = DisplayRound(state, -balance.NetOverride.Value); result.Source = "人工填写的单项工程净核减总额\r\n" + result.Source; }
            else if (totals.Count == 0) result.Issue = "没有对应计算表，请导入包含计算表的文件或填写总额";
            else if (totals.Any(t => !t.Delta.HasValue || t.Message != "")) result.Issue = string.Join("；", totals.Where(t => !t.Delta.HasValue || t.Message != "").Select(t => t.Sheet + "：" + t.Message));
            else if (state.Items.Any(i => totals.Any(t => t.Key == i.CalculationKey) && Norm(i.Group) != Norm(balance.Group))) result.Issue = "关联明细的单项工程名称已改变，请核对总额归属后设置";
            else if (totals.Any(t => t.DetailSheets == null || t.DetailSheets.Count == 0) || items.Any(i => string.IsNullOrEmpty(i.CalculationKey) || !totals.Any(t => t.Key == i.CalculationKey))) result.Issue = "计算表与明细对应不完整，请核对工作表编号及工程名称或填写核实后的总额";
            else result.TotalDelta = DisplayRound(state, totals.Sum(t => t.Delta.Value));
            if (result.TotalDelta.HasValue) { result.Residual = DisplayRound(state, result.TotalDelta.Value - result.Listed); result.NonDetail = DisplayRound(state, result.Residual - result.Omitted); }
            if (!Targets.Contains(balance.Target)) result.Issue = "请选择差额说明的报告归属";
            if (string.IsNullOrWhiteSpace(balance.Group) || balance.Group == "未填写工程名称") result.Issue = "计算表缺少单项工程名称，请核对表头后重新导入";
            return result;
        }
        public static List<string> BalanceIssues(ExcelReviewState state) { Normalize(state); return state.Balances.Where(b => b.Enabled).Select(b => new { b.Group, Result = CalculateBalance(state, b) }).Where(x => x.Result.Issue != "").Select(x => x.Group + "：" + x.Result.Issue).ToList(); }
    }
}
