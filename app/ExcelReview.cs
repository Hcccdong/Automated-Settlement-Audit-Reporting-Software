using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace SettlementReview {
    public static partial class ExcelReview {
        public static readonly string[] Reasons = { "工程量偏差", "单价偏差", "工程量及单价偏差", "其他原因（填写说明）", "待确认" };
        public static readonly string[] Targets = { "合同标内调整说明", "合同标外调整说明", "联系单调整说明", "其他调整说明" };
        public static readonly string[] Categories = { "分部分项", "技术措施", "待分类" };
        public const string DirectionWording = "正负偏差";
        public const string NetWording = "净核减增";
        public static string DetectCategory(string title) { string t = Norm(title); return t.Contains("技术措施") || t.Contains("技措") ? Categories[1] : t.Contains("分部分项") || t.Contains("分部工程") ? Categories[0] : Categories[2]; }
        public static string DetectTarget(string title) { return title.Contains("合同标外") ? Targets[1] : title.Contains("联系单") ? Targets[2] : Targets[0]; }
        public static string CleanGroup(string value) { return Regex.Replace(value ?? "", @"\s*第\s*\d+\s*页.*$", "").Trim(); }
        public static string ProjectTitle(ExcelSheet sheet, int row) {
            string joined = string.Join(" ", sheet.Rows[row].Where(c => c.Key <= 8).OrderBy(c => c.Key).Select(c => c.Value));
            var match = Regex.Match(joined, @"^\s*工\s*程\s*名\s*称\s*[:：]?\s*(?<name>.+)$");
            return match.Success ? CleanGroup(match.Groups["name"].Value) : "";
        }
        public static string Norm(string s) { return Regex.Replace(s ?? "", @"\s+", "").Replace("（", "(").Replace("）", ")"); }
        public static string FullText(string s) { return Regex.Replace((s ?? "").Replace("\\n", "；").Replace("\r", "").Replace("\n", "；"), @"[ \t]+", " ").Trim(); }
        public static string Number(decimal? d) { return d.HasValue ? d.Value.ToString("0.########", CultureInfo.InvariantCulture) : "未提供"; }
        public static string Money(decimal? d) { return d.HasValue ? d.Value.ToString("0.00", CultureInfo.InvariantCulture) : "未提供"; }
        public static decimal DisplayRound(ExcelReviewState state, decimal value) { Normalize(state); return decimal.Round(value, state.AmountDecimals, MidpointRounding.AwayFromZero); }
        public static string Money(ExcelReviewState state, decimal? d) { return d.HasValue ? DisplayRound(state, d.Value).ToString("F" + state.AmountDecimals, CultureInfo.InvariantCulture) : "未提供"; }
        public static string DirectionText(ExcelReviewState state, ReviewItem item) {
            if (item == null || !item.Delta.HasValue) return "待确认";
            if (item.Delta.Value == 0) return "无偏差";
            Normalize(state);
            return state.WordingMode == NetWording ? (item.Delta.Value > 0 ? "净核增" : "净核减") : (item.Delta.Value > 0 ? "正偏差" : "负偏差");
        }
        public static List<ExcelSheet> Read(string path) {
            return ExcelWorkbookReader.Read(path);
        }
        static readonly Dictionary<string, string[]> Aliases = new Dictionary<string, string[]> {
            { "name", new[] { "项目名称", "清单名称", "名称", "工程名称" } }, { "code", new[] { "项目编码", "清单编码", "编号", "编码" } },
            { "feature", new[] { "项目特征描述", "项目特征", "特征描述", "规格型号" } }, { "unit", new[] { "计量单位", "单位" } }, { "group", new[] { "单位工程", "单位工程名称" } },
            { "sq", new[] { "送审工程量", "送审数量" } }, { "aq", new[] { "审查工程量", "审查数量", "审核工程量", "审核数量", "审定工程量" } },
            { "sp", new[] { "送审综合单价", "送审单价" } }, { "ap", new[] { "审查综合单价", "审查单价", "审核综合单价", "审核单价", "审定单价" } },
            { "sa", new[] { "施工方送审金额", "送审金额", "送审合价", "送审造价" } }, { "aa", new[] { "审查金额", "审查合价", "审核金额", "审核合价", "审定金额", "审核造价" } },
            { "inc", new[] { "核增", "核增金额", "核增额" } }, { "dec", new[] { "核减", "核减金额", "核减额" } }, { "net", new[] { "净核减", "净核减金额" } }
        };
        static Dictionary<string, int> Columns(ExcelSheet s, int row) {
            var map = new Dictionary<string, int>(); string side = "";
            for (int col = 1; col <= 150; col++) {
                string text = Norm(s.Cell(row, col)), parent = Norm(s.Cell(row - 1, col, true));
                if (parent == "送审") side = "s"; else if (parent == "审查" || parent == "审核" || parent == "审定") side = "a";
                if (text == "" && (parent == "核增" || parent == "核减" || parent == "序号")) text = parent;
                if (text == "序号") { map["no"] = col; continue; }
                text = Regex.Replace(text, @"\((?:元|万元|m2|m3)\)$", "");
                if (side != "" && new[] { "编号", "编码", "名称", "项目名称", "项目特征", "项目特征描述", "单位", "计量单位", "数量", "工程量", "单价", "综合单价", "合价", "金额" }.Contains(text)) {
                    string key = text.Contains("特征") ? "feature" : text.Contains("名称") ? "name" : text.Contains("编号") || text.Contains("编码") ? "code" : text.Contains("单位") ? "unit" : text == "数量" || text == "工程量" ? "q" : text.Contains("单价") ? "p" : "a";
                    if (new[] { "name", "code", "feature", "unit" }.Contains(key)) key = side == "s" ? key : "audit_" + key; else key = side + key;
                    if (!map.ContainsKey(key)) map[key] = col;
                } else foreach (var pair in Aliases) if (pair.Value.Contains(text) && !map.ContainsKey(pair.Key)) { map[pair.Key] = col; break; }
            }
            return map;
        }
        public static ExcelReviewState Parse(List<ExcelSheet> sheets, string name) {
            var state = new ExcelReviewState { FileName = name, CompactFeatures = true, ModelVersion = 11, WordingMode = DirectionWording, AmountDecimals = 2, Totals = ReadTotals(sheets), Balances = new List<ReviewBalance>() };
            foreach (var s in sheets) {
                if (IsCalculation(s) || Norm(s.Name).Contains("工程项目结算审核对比表")) continue;
                Dictionary<string, int> map = null; int header = 0;
                foreach (int row in s.Rows.Keys.Where(r => r <= 40)) { var candidate = Columns(s, row); if (candidate.ContainsKey("name") && (candidate.ContainsKey("sa") && candidate.ContainsKey("aa") || candidate.ContainsKey("inc") || candidate.ContainsKey("dec") || candidate.ContainsKey("net")) && (map == null || candidate.Count > map.Count)) { map = candidate; header = row; } }
                if (map == null) { state.Messages.Add(s.Name + "：未识别到对比表表头，未计入。"); continue; }
                bool tenThousands = s.Rows.Where(r => r.Key <= header).SelectMany(r => r.Value.Values).Any(v => Norm(v).Contains("万元"));
                if (tenThousands) { state.Messages.Add(s.Name + "：表头含万元口径，本版要求金额为元，请核对并换算后导入；本表未计入。"); continue; }
                string group = "未填写工程名称"; bool hasProjectTitle = false; int priorCount = state.Items.Count;
                string category = DetectCategory(s.Name);
                if (category == Categories[2]) category = s.Rows.Where(r => r.Key <= header).SelectMany(r => r.Value.Values).Select(DetectCategory).FirstOrDefault(c => c != Categories[2]) ?? Categories[2];
                foreach (int r in s.Rows.Keys) {
                    string title = ProjectTitle(s, r); if (title != "") { group = title; hasProjectTitle = true; continue; }
                    if (r <= header) continue;
                    string section = s.Rows[r].Values.Select(v => Regex.Replace(Norm(v), @"^[一二三四五六七八九十0-9、.（）()]+", "")).FirstOrDefault(v => new[] { "分部分项", "分部分项工程", "分部分项工程费", "技术措施", "技术措施项目", "技术措施项目费", "施工技术措施项目费" }.Contains(v));
                    if (section != null) { category = DetectCategory(section); continue; }
                    Func<string, string> get = key => map.ContainsKey(key) ? s.Cell(r, map[key]).Trim() : "";
                    if (map.ContainsKey("audit_name") && map.ContainsKey("no") && !Regex.IsMatch(get("no"), @"^\d+(?:\.0+)?$")) continue;
                    string itemName = get("name"), auditName = get("audit_name");
                    if (itemName == "" && auditName == "") continue;
                    if (new[] { "名称", "项目名称", "合计", "小计", "总计", "本页小计", "累计", "本页合计" }.Contains(Norm(itemName)) || Norm(itemName).Contains("合计") || Norm(itemName).Contains("总计")) continue;
                    if (Norm(itemName).Contains("简明表")) continue;
                    if (get("group") != "" && !hasProjectTitle) group = CleanGroup(get("group"));
                    if (new[] { "sq", "aq", "sp", "ap", "sa", "aa", "inc", "dec", "net" }.All(k => get(k) == "")) continue;
                    var item = new ReviewItem { Sheet = s.Name, Row = r, Group = group, Category = category, Target = DetectTarget(s.Name), Code = get("code"), Name = itemName, Feature = get("feature"), Unit = get("unit"), AuditCode = get("audit_code"), AuditName = auditName, AuditFeature = get("audit_feature"), AuditUnit = get("audit_unit"), Note = "", Warning = "", Reason = "待确认" };
                    var warnings = new List<string>();
                    Func<string, decimal?> number = key => {
                        string raw = get(key); if (raw == "" || raw == "-" || raw == "—") return null;
                        decimal result; if (decimal.TryParse(raw.Replace(",", "").Replace("，", ""), NumberStyles.Float | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out result)) return result;
                        warnings.Add(key + " 列数值不可用（" + raw + "）"); return null;
                    };
                    item.SentQuantity = number("sq"); item.AuditQuantity = number("aq"); item.SentPrice = number("sp"); item.AuditPrice = number("ap"); item.SentAmount = number("sa"); item.AuditAmount = number("aa"); item.SourceIncrease = number("inc"); item.SourceDecrease = number("dec"); decimal? net = number("net");
                    if (item.SentAmount.HasValue && item.AuditAmount.HasValue) item.Delta = Round(item.AuditAmount.Value - item.SentAmount.Value);
                    else if (item.SourceIncrease.HasValue || item.SourceDecrease.HasValue) item.Delta = Round((item.SourceIncrease ?? 0) - (item.SourceDecrease ?? 0));
                    else if (net.HasValue) item.Delta = -Round(net.Value);
                    if (!item.Delta.HasValue) warnings.Add("缺少可用的送审/审查金额及偏差金额");
                    if (item.SourceIncrease < 0 || item.SourceDecrease < 0) warnings.Add("核增、核减栏目应填非负金额，请核实口径");
                    if (item.Delta.HasValue && (item.SourceIncrease.HasValue || item.SourceDecrease.HasValue) && Round((item.SourceIncrease ?? 0) - (item.SourceDecrease ?? 0)) != item.Delta) warnings.Add("核增减列与审查减送审金额不一致，组稿按审查减送审计算");
                    if (item.Delta.HasValue && net.HasValue && -Round(net.Value) != item.Delta) warnings.Add("净核减列与送审减审查金额不一致");
                    foreach (var pair in new[] { new[] { item.Code, item.AuditCode, "编码" }, new[] { item.Name, item.AuditName, "名称" }, new[] { item.Feature, item.AuditFeature, "项目特征" }, new[] { item.Unit, item.AuditUnit, "计量单位" } })
                        if (map.ContainsKey("audit_name") && Norm(pair[0]) != Norm(pair[1])) warnings.Add("送审/审查" + pair[2] + "不一致，请核对是否同一项");
                    item.Reason = InferReason(item); item.Warning = string.Join("；", warnings.Distinct());
                    item.Include = item.Delta.HasValue && item.Delta != 0 && item.Warning == "";
                    state.Items.Add(item); if (state.Items.Count > 10000) throw new InvalidDataException("对比明细超过10000条，请分批导入。");
                }
                state.Messages.Add(s.Name + "：识别 " + (state.Items.Count - priorCount) + " 条明细（含无偏差项，汇总行已排除）。");
            }
            if (state.Items.Count == 0) state.Messages.Add("没有可用的对比明细。请检查表头及已填写的送审、审查金额。");
            LinkTotals(state, sheets);
            return state;
        }
        static decimal Round(decimal value) { return decimal.Round(value, 8, MidpointRounding.AwayFromZero); }
        public static string InferReason(ReviewItem i) {
            if (i.Unit != "" && i.AuditUnit != "" && Norm(i.Unit) != Norm(i.AuditUnit)) return "待确认";
            if (!i.SentQuantity.HasValue || !i.AuditQuantity.HasValue || !i.SentPrice.HasValue || !i.AuditPrice.HasValue) return "待确认";
            bool quantity = i.SentQuantity != i.AuditQuantity, price = i.SentPrice != i.AuditPrice;
            return quantity && price ? Reasons[2] : quantity ? Reasons[0] : price ? Reasons[1] : Reasons[4];
        }
        public static List<ReviewItem> Selected(ExcelReviewState state) { return state.Items.Where(i => i.Include && i.Delta.HasValue && i.Delta != 0 && Math.Abs(i.Delta.Value) >= state.Threshold).ToList(); }
        public static string Summary(ExcelReviewState state) {
            var rows = Selected(state); decimal inc = rows.Where(i => i.Delta > 0).Sum(i => i.Delta.Value), dec = -rows.Where(i => i.Delta < 0).Sum(i => i.Delta.Value);
            return "采用 " + rows.Count + " 项  ·  核增 " + Money(state, inc) + " 元  ·  核减 " + Money(state, dec) + " 元  ·  净核减 " + Money(state, dec - inc) + " 元\r\n所选明细合计，不自动替代报告总额。";
        }
        public static string Describe(string name, string feature) {
            string title = FullText(name).TrimEnd('。', '；', '，'); var pieces = new List<string>();
            foreach (string part in FullText(feature).Split(new[] { '；' }, StringSplitOptions.RemoveEmptyEntries)) {
                string text = Regex.Replace(part.Trim(), @"^\s*\d+(?:[、．]|\.(?!\d))\s*", "").TrimEnd('。', '，', '；');
                text = Regex.Replace(text, @"^项目特征(?:描述)?\s*[:：]?\s*", "");
                if (Norm(text) != "" && Norm(text) != Norm(title) && !pieces.Any(p => Norm(p) == Norm(text))) pieces.Add(text);
            }
            return title + (pieces.Count == 0 ? "" : "，" + string.Join("，", pieces));
        }
        public static void Normalize(ExcelReviewState state) {
            if (state.ModelVersion < 9) state.CompactFeatures = true;
            if (state.ModelVersion < 10 || state.WordingMode != DirectionWording && state.WordingMode != NetWording) state.WordingMode = DirectionWording;
            if (state.ModelVersion < 11 || state.AmountDecimals < 0 || state.AmountDecimals > 4) state.AmountDecimals = 2;
            state.ModelVersion = 11;
            if (state.Totals == null) state.Totals = new List<ReviewTotal>(); if (state.Balances == null) state.Balances = new List<ReviewBalance>();
            foreach (var item in state.Items) { if (!Categories.Contains(item.Category)) item.Category = DetectCategory(item.Sheet ?? ""); if (!Targets.Contains(item.Target)) item.Target = DetectTarget(item.Sheet ?? ""); item.Group = CleanGroup(item.Group); }
            ReviewGroupEditor.RepairAssociations(state);
        }
        public static string Compose(ExcelReviewState state, bool strict, string target = null) {
            Normalize(state); var lines = new List<string>(); int number = 0;
            var descriptors = BuildDescriptions(state);
            var selected = Selected(state).Where(i => target == null || i.Target == target).ToList();
            var groupNames = selected.Select(i => i.Group).Concat(state.Balances.Where(b => b.Enabled && (target == null || b.Target == target)).Select(b => b.Group)).GroupBy(Norm).Select(g => g.First()).ToList();
            foreach (var groupName in groupNames) {
                var group = selected.Where(i => Norm(i.Group) == Norm(groupName));
                int headingIndex = lines.Count;
                lines.Add(groupName + "：");
                foreach (var i in group) {
                if (strict && RowIssue(i) != "") throw new InvalidDataException(i.Location + "：" + RowIssue(i));
                string identity = descriptors[i];
                string values = "送审金额为" + Money(state, i.SentAmount) + (i.SentAmount.HasValue ? "元" : "") + "，审核金额为" + Money(state, i.AuditAmount) + (i.AuditAmount.HasValue ? "元" : "");
                if (state.Detailed) values = "送审工程量" + Number(i.SentQuantity) + i.Unit + "、单价" + Number(i.SentPrice) + "元；审查工程量" + Number(i.AuditQuantity) + (i.AuditUnit == "" ? i.Unit : i.AuditUnit) + "、单价" + Number(i.AuditPrice) + "元；" + values;
                string reason = i.Reason == Reasons[3] ? (string.IsNullOrWhiteSpace(i.Note) ? "调整原因说明待填写" : i.Note.Trim()) : !Reasons.Contains(i.Reason) || i.Reason == "待确认" ? "调整原因待确认" : "由于" + i.Reason + "，审核予以调整";
                if (i.Reason != Reasons[3] && !string.IsNullOrWhiteSpace(i.Note)) reason += "（" + i.Note.Trim() + "）";
                lines.Add((++number) + ". " + identity + "，" + values + "，" + reason + "，" + DeviationText(state, i.Delta.Value, false) + "。" + (i.Warning != "" && !i.Confirmed ? "【数据提示待核实】" : ""));
            }
                var balance = state.Balances.FirstOrDefault(b => b.Enabled && Norm(b.Group) == Norm(groupName) && (target == null || b.Target == target));
                if (balance != null) {
                    var result = CalculateBalance(state, balance);
                    if (strict && result.Issue != "") throw new InvalidDataException(groupName + "：" + result.Issue);
                    if (result.Issue != "") lines.Add("【总额差额待核对：" + result.Issue + "】");
                    else if (result.Residual != 0) lines.Add((++number) + ". " + (string.IsNullOrWhiteSpace(balance.Reason) ? FeeReason : balance.Reason.Trim().TrimEnd('。', '，', '；')) + "，" + DeviationText(state, result.Residual, true) + "。");
                }
                if (lines.Count == headingIndex + 1) lines.RemoveAt(headingIndex);
            }
            if (strict && lines.Count == 0 && target == null) throw new InvalidDataException("请先选择有偏差的明细，再生成说明。");
            return lines.Count == 0 && target != null ? "无" : string.Join("\r\n", lines);
        }
        public static string RowIssue(ReviewItem item) {
            var issues = new List<string>();
            if (!item.Delta.HasValue) issues.Add("缺少可用的偏差金额");
            if (!Reasons.Contains(item.Reason) || item.Reason == "待确认") issues.Add("请选择调整原因");
            if (item.Reason == Reasons[3] && string.IsNullOrWhiteSpace(item.Note)) issues.Add("其他原因需填写补充说明");
            if (!string.IsNullOrWhiteSpace(item.Warning) && !item.Confirmed) issues.Add("配对或金额提示未核实，请核对后勾选“提示已核实”或取消采用");
            if (string.IsNullOrWhiteSpace(item.Group) || item.Group == "未填写工程名称") issues.Add("请填写单项工程名称");
            return string.Join("；", issues);
        }
        public static List<ReviewItem> ProblemRows(ExcelReviewState state) {
            return state == null ? new List<ReviewItem>() : state.Items.Where(i => i.Include && (!i.Delta.HasValue || i.Delta != 0 && Math.Abs(i.Delta.Value) >= state.Threshold) && RowIssue(i) != "").ToList();
        }
        public static void StartLink(Project project, ExcelReviewState state) {
            Normalize(state);
            if (state.OriginalBodies == null) state.OriginalBodies = Targets.ToDictionary(t => t, t => Core.Get(project, t));
            state.AutoLink = true; state.ModelVersion = 11; project.Excel = state; Synchronize(project, false);
        }
        public static void Synchronize(Project project, bool strict) {
            var state = project.Excel; if (state == null || !state.AutoLink) return;
            if (strict) { var problem = ProblemRows(state).FirstOrDefault(); if (problem != null) throw new InvalidDataException(problem.Location + "：" + RowIssue(problem)); }
            var texts = Targets.ToDictionary(t => t, t => Compose(state, strict, t));
            foreach (var pair in texts) project.Values[pair.Key] = pair.Value;
        }
        public static void Apply(Project project, string target, string text) {
            if (!Targets.Contains(target)) throw new InvalidDataException("请选择调整说明的填写位置。");
            string existing = Core.Get(project, target); var state = project.Excel;
            if (state.LastTarget == target && state.LastText != "" && existing.Contains(state.LastText)) existing = existing.Replace(state.LastText, text);
            else if (!existing.Contains(text)) existing = existing == "" || existing == "无" ? text : existing + "\r\n\r\n" + text;
            project.Values[target] = existing; state.LastTarget = target; state.LastText = text;
        }
    }
}
