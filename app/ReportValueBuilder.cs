using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace SettlementReview {
    // Application orchestration for template values; retains synchronization order.
    public static class ReportValueBuilder {
        public static Dictionary<string, string> Build(Project p, bool draft, IEnumerable<string> required = null) {
            ExcelReview.Synchronize(p, !draft);
            if (!draft && Core.Get(p, "下浮率模式") == "自动" && Core.DiscountInputs.Any(k => Core.Get(p, k) != "")) p.Values["工程下浮率"] = Core.DiscountText(Core.DiscountRate(p));
            var d = Core.Fields.ToDictionary(f => f.Name, f => Core.Get(p, f.Name));
            d["审核依据及原则"] = Core.ProjectBasis(p);
            string reportDateText = d["报告日期"]; var reportDate = Core.ParseDate(reportDateText);
            if (reportDateText != "" && !reportDate.HasValue) throw new InvalidDataException("报告日期无效，请使用日期选择框。");
            if (reportDate.HasValue) d["报告日期"] = reportDate.Value.ToString("yyyy年M月d日");
            if (d["净核减金额"] == "" && d["送审金额"] != "" && d["审定金额"] != "") d["净核减金额"] = Core.F(Core.Money(d["送审金额"], "送审金额") - Core.Money(d["审定金额"], "审核金额"));
            if (!draft) {
                string[] optional = { "审核人员", "复核人员", "工期补充说明" };
                var requiredNames = required == null ? new HashSet<string>(Core.Fields.Select(f => f.Name)) : new HashSet<string>(required);
                var missing = Core.Fields.Where(f => requiredNames.Contains(f.Name) && !optional.Contains(f.Name) && string.IsNullOrWhiteSpace(d[f.Name])).Select(f => f.Name).ToArray();
                if (missing.Length > 0) throw new InvalidDataException("请补齐以下内容（无对应事项填写“无”）：\n" + string.Join("、", missing.Select(k => Core.Fields.First(f => f.Name == k).Label)));
            }
            foreach (var key in new[] { "合同金额", "送审金额", "审定金额", "核减合计", "审增合计" }) {
                if (d[key] != "" && !(key == "合同金额" && Core.IsAbsent(d[key]))) d[key] = Core.F(Core.Money(d[key], key));
            }
            bool complete = new[] { "送审金额", "审定金额", "核减合计", "审增合计" }.All(k => d[k] != "");
            if (complete) {
                decimal before = Core.Money(d["送审金额"], "送审金额"), after = Core.Money(d["审定金额"], "审定金额");
                decimal reduction = Core.Money(d["核减合计"], "核减合计"), increase = Core.Money(d["审增合计"], "审增合计"), net = before - after;
                d["净核减金额"] = Core.Get(p, "净核减金额") == "" ? Core.F(net) : Core.F(Core.SignedMoney(Core.Get(p, "净核减金额")));
                d["净核减率"] = before == 0 ? "不适用（送审金额为零）" : (net / before).ToString("0.00%", CultureInfo.InvariantCulture);
            } else { d["净核减金额"] = Core.Get(p, "净核减金额") == "" ? "待填写" : Core.F(Core.SignedMoney(Core.Get(p, "净核减金额"))); d["净核减率"] = "待填写"; }
            foreach (var key in d.Keys.ToArray()) {
                if (d[key].Contains("{{") || d[key].Contains("}}")) throw new InvalidDataException(key + "：请填写实际内容，不要保留模板字段标记。");
                if (string.IsNullOrWhiteSpace(d[key])) d[key] = key == "审核人员" || key == "复核人员" ? "________________" : "待填写";
            }
            foreach (var fixedText in Core.FixedBody) if (fixedText.Key != "审核依据及原则") d[fixedText.Key] = fixedText.Value;
            foreach (var schedule in Core.ScheduleValues(p, draft)) d[schedule.Key] = schedule.Value;
            d["有关事项说明"] = string.Join("\n", Core.MatterFields.Where(k => !Core.IsAbsent(d[k])).Select(k => k + "：" + d[k]));
            d["抄送单位"] = string.Join("、", new[] { "建设单位", "施工单位", "代建单位" }.Select(k => Core.Get(p, k)).Where(s => s != "" && s != "无" && s != "不适用" && s != "无代建单位").Distinct());
            d["审核金额"] = d["审定金额"]; d["核减金额"] = d["核减合计"]; d["核增金额"] = d["审增合计"];
            foreach (string hidden in new[] { "审核人员", "复核人员", "税费口径", "审核结论补充说明" }) d[hidden] = "";
            d["金额比较口径"] = "建设单位审核价";
            decimal signedNet; bool validNet = decimal.TryParse(d["净核减金额"], NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out signedNet);
            d["净额方向"] = validNet && signedNet < 0 ? "净核增" : "净核减"; d["净额绝对金额"] = validNet ? Core.F(Math.Abs(signedNet)) : d["净核减金额"];
            d["金额汇总说明"] = "本工程建设单位审核价为" + d["送审金额"] + "元，审核金额为" + d["审定金额"] + "元，核增" + d["审增合计"] + "元，核减" + d["核减合计"] + "元，" + d["净额方向"] + d["净额绝对金额"] + "元。修正情况分析如下：";
            d["审核结论"] = "本工程建设单位审定金额为" + d["送审金额"] + "元；审核金额为" + d["审定金额"] + "元；核增金额为" + d["审增合计"] + "元；核减金额为" + d["核减合计"] + "元；" + d["净额方向"] + "金额为" + d["净额绝对金额"] + "元。";
            d["修正造价情况说明"] = string.Join("\n", ExcelReview.Targets.Where(k => !Core.IsAbsent(d[k])).Select(k => k.Replace("说明", "") + "：\n" + d[k]));
            d["签发日期"] = d["报告日期"];
            return d;
        }
    }
}
