using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace SettlementReview {
    public static partial class Core {
        public static void RemoveAutomaticNumberSuffix(XDocument xml) {
            foreach (var paragraph in xml.Descendants(W + "p").ToList()) {
                string text = RawParagraphText(paragraph);
                var match = Regex.Match(text, @"^\s*\{\{报告编号\}\}(?<suffix>\s*[-－—]\s*\{\{项目名称\}\})\s*$");
                if (match.Success) ReplaceSpan(paragraph, match.Groups["suffix"].Index, match.Groups["suffix"].Length, "");
            }
        }
    }
}
