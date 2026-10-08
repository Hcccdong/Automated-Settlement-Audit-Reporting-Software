using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SettlementReview {
    public static partial class Tests {
        static void CheckVersionTwelve(string output, Project example, byte[] template) {
            var unique = ExcelReview.Parse(new List<ExcelSheet> { ReviewTestSheet("明细", "1号楼", "挖基坑土方", "土壤类别详见地质报告及现场踏勘，机械开挖，人工配合", "161686", "154908") }, "唯一名称.xlsx");
            string uniqueText = ExcelReview.Compose(unique, true);
            Assert(uniqueText.Contains("挖基坑土方，送审金额为") && !uniqueText.Contains("土壤类别") && !uniqueText.Contains("机械开挖"), "精简模式下名称无歧义时完全省略项目特征");

            var duplicateSheet = ReviewTestSheet("明细", "1号楼", "砖基础", "基础砖与工艺按图，MU10水泥标准砖，M7.5水泥砂浆", "1000", "900");
            ExcelRow(duplicateSheet, 5, "砖基础", "基础砖与工艺按图，MU15水泥标准砖，M7.5水泥砂浆", "10", "8", "100", "100", "1000", "850");
            var duplicate = ExcelReview.Parse(new List<ExcelSheet> { duplicateSheet }, "同名名称.xlsx"); var descriptions = ExcelReview.BuildDescriptions(duplicate);
            Assert(descriptions[duplicate.Items[0]].Contains("MU10") && descriptions[duplicate.Items[1]].Contains("MU15") && !descriptions[duplicate.Items[0]].Contains("基础砖与工艺") && !descriptions[duplicate.Items[0]].Contains("M7.5"), "同名项目只补充一项最短区别特征");
            Assert(ExcelReview.Compose(duplicate, true).Length < 500, "同名项目精简后不带入整段共同项目特征");
            duplicate.CompactFeatures = false; Assert(ExcelReview.Compose(duplicate, true).Contains("基础砖与工艺按图") && ExcelReview.Compose(duplicate, true).Contains("M7.5水泥砂浆"), "关闭精简模式仍可输出完整特征");

            byte[] docx = Core.Export(example, template, true);
            using (var preview = new RichTextBox()) {
                WordPreviewRenderer.Render(preview, docx);
                Assert(preview.Text.Contains("结算审核报告") && preview.Text.Contains(example.Values["项目名称"]) && !preview.Text.Contains("草稿  待审核确认"), "Word样式预览显示最终报告正文且隐藏草稿标记");
                int title = preview.Text.IndexOf("结算审核报告"); preview.Select(title, "结算审核报告".Length);
                Assert(preview.SelectionAlignment == HorizontalAlignment.Center && preview.SelectionFont != null && preview.SelectionFont.Bold && preview.SelectionFont.Size >= 15F, "Word样式预览保留标题居中、加粗和字号");
                Assert(preview.Rtf.Contains("\\fonttbl") && preview.Rtf.Contains("\\qc"), "Word样式预览生成富文本字体表和段落对齐");
            }

            string isolated = Path.Combine(output, "v0120_ui_data", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(isolated);
            using (var form = new MainForm(true, isolated)) {
                V7Call(form, "LoadIntoForm", example); form.Show(); Application.DoEvents();
                var tabs = V7Field<TabControl>(form, "tabs"); tabs.SelectedIndex = 6; Application.DoEvents(); V7Call(form, "RefreshPreview"); Application.DoEvents();
                var rich = V7Field<RichTextBox>(form, "preview");
                Assert(rich.Text.Contains("结算审核报告") && rich.ReadOnly && rich.BackColor == Color.White, "报告导出页使用只读白色Word样式预览区");
                Assert(ControlsBelow(form.Controls).OfType<Button>().Any(b => b.Text == "刷新 Word 预览"), "报告导出页提供刷新Word预览按钮");
                form.Screenshot(Path.Combine(output, "v0120_Word样式预览.png"), 6); form.Close();
            }
        }
    }
}
