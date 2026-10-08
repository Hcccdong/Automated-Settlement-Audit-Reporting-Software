using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace SettlementReview {
    public static partial class ExcelReview {
        public static string DetectWordingMode(byte[] template) {
            string text = string.Join("\n", Core.TemplateParagraphs(template).Select(p => p.Text));
            int direction = Regex.Matches(text, @"(?:正|负)偏差").Count;
            int net = Regex.Matches(text, @"净核(?:减|增)").Count;
            return direction > 0 ? DirectionWording : net > 0 ? NetWording : DirectionWording;
        }
        public static string WordingLabel(ExcelReviewState state) { Normalize(state); return state.WordingMode == NetWording ? "净核减 / 净核增" : "正偏差 / 负偏差"; }
        public static string DeviationText(ExcelReviewState state, decimal delta, bool balance) {
            Normalize(state); string money = Money(state, Math.Abs(delta)) + "元";
            if (state.WordingMode == NetWording) return (delta > 0 ? "净核增" : "净核减") + money;
            return "此项" + (delta > 0 ? "正偏差" : "负偏差") + (balance ? "为" : "") + money;
        }
        static List<string> FeatureParts(string name, string feature) {
            var result = new List<string>();
            string source = (feature ?? "").Replace("\\n", "\n").Replace("\r", "");
            // A sequence on separate lines is a list; an isolated 1.5 is a specification.
            var numbered = Regex.Matches(source, @"(?m)^\s*(\d+)\.(.+)$").Cast<Match>().ToList();
            bool list = numbered.Count >= 2 && numbered[0].Groups[1].Value == "1" && numbered[1].Groups[1].Value == "2" && !char.IsDigit(numbered[0].Groups[2].Value[0]);
            if (list) source = Regex.Replace(source, @"(?m)^\s*\d+\.\s*", "");
            source = Regex.Replace(source, @"[；;]\s*([（(](?:另列|另计|不含)[^）)]*[）)])", "$1");
            foreach (string part in FullText(source).Split(new[] { '；' }, StringSplitOptions.RemoveEmptyEntries).SelectMany(p => Regex.Split(p, @"[，,]")).Where(p => !string.IsNullOrWhiteSpace(p))) {
                string text = Regex.Replace(part.Trim(), @"^\s*\d+(?:[、．]|\.(?!\d))\s*", "").TrimEnd('。', '，', '；');
                text = Regex.Replace(text, @"^(?:项目特征(?:描述)?|混凝土强度等级|砂浆强度等级|材料名称|规格型号)\s*[:：]\s*", "");
                if (Norm(text) == "" || Norm(name).Contains(Norm(text))) continue;
                if (Regex.IsMatch(text, @"^(?:其他未尽事宜|未尽事宜|具体详见|详见设计图纸|满足设计及规范要求|符合设计及规范要求|工作内容详见|施工要求详见)")) continue;
                if (!result.Any(p => Norm(p) == Norm(text))) result.Add(text);
            }
            return result;
        }
        static string ShortDescription(string name, string feature, List<string> peers) {
            string title = FullText(name).TrimEnd('。', '；', '，');
            var distinct = peers.Select(f => Norm(f)).Distinct().ToList();
            if (distinct.Count <= 1) return title;
            var parts = FeatureParts(name, feature);
            if (parts.Count == 0) return title;
            var otherParts = peers.Where(f => Norm(f) != Norm(feature)).Select(f => FeatureParts(name, f)).ToList();
            var distinctive = parts.Where(p => otherParts.Any(other => !other.Any(o => Norm(o) == Norm(p)))).ToList();
            if (distinctive.Count == 0) return title;
            string chosen = distinctive.OrderBy(p => Regex.IsMatch(p, @"\d|[Cc][0-9]|强度|厚|规格|材质|砂浆|砌块|防水|木模|铝模") ? 0 : 1).ThenBy(p => p.Length).First();
            return title + "，" + chosen;
        }
        public static Dictionary<ReviewItem, string> BuildDescriptions(ExcelReviewState state) {
            var peers = new Dictionary<string, List<string>>();
            foreach (var item in state.Items) foreach (var pair in new[] { new[] { item.Name, item.Feature }, new[] { item.AuditName, item.AuditFeature } }) {
                string key = Norm(item.Group) + "\0" + Norm(pair[0]); if (!peers.ContainsKey(key)) peers[key] = new List<string>(); peers[key].Add(pair[1] ?? "");
            }
            var cache = new Dictionary<string, string>();
            Func<ReviewItem, string, string, string> describe = (item, name, feature) => {
                string key = Norm(item.Group) + "\0" + Norm(name), cacheKey = key + "\0" + (feature ?? "");
                if (!cache.ContainsKey(cacheKey)) cache[cacheKey] = state.CompactFeatures ? ShortDescription(name, feature, peers.ContainsKey(key) ? peers[key] : new List<string>()) : Describe(name, feature);
                return cache[cacheKey];
            };
            var result = new Dictionary<ReviewItem, string>();
            foreach (var item in state.Items) {
                string name = string.IsNullOrEmpty(item.Name) ? item.AuditName : item.Name, feature = string.IsNullOrEmpty(item.Feature) ? item.AuditFeature : item.Feature;
                result[item] = describe(item, name, feature);
                if (!string.IsNullOrEmpty(item.AuditName) && (Norm(item.Name) != Norm(item.AuditName) || Norm(item.Feature) != Norm(item.AuditFeature))) result[item] = "送审为" + describe(item, item.Name, item.Feature) + "，审核为" + describe(item, item.AuditName, item.AuditFeature);
            }
            return result;
        }
    }
}
