using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;

namespace SettlementReview {
    public static partial class Tests {
        static void CheckVersionTen(string output, Project example, byte[] template) {
            byte[] source = Core.SourceTemplate();
            Assert(ExcelReview.DetectWordingMode(source) == ExcelReview.DirectionWording, "默认模板按调整明细识别为正负偏差口径");
            string xml = Read(source, "word/document.xml").Replace("负偏差", "净核减").Replace("正偏差", "净核增");
            byte[] netTemplate = Change(source, "word/document.xml", xml);
            Assert(ExcelReview.DetectWordingMode(netTemplate) == ExcelReview.NetWording, "仅含净核减增示例的模板识别为净额口径");
            var state = V9Example(); state.Threshold = 500; state.WordingMode = ExcelReview.NetWording;
            string netText = ExcelReview.Compose(state, true);
            Assert(netText.Contains("净核减9000.00元") && netText.Contains("净核减1000.00元") && !netText.Contains("负偏差"), "净额口径同时作用于明细和单项工程末条配平");
            state.Items[0].AuditAmount = 24000; state.Items[0].Delta = 9000; state.Balances[0].NetOverride = -9100;
            Assert(ExcelReview.Compose(state, true).Contains("净核增9000.00元") && ExcelReview.Compose(state, true).Contains("净核增100.00元"), "正向金额在净额口径中写为净核增");
            state.WordingMode = ExcelReview.DirectionWording; state.Items[0].AuditAmount = 6000; state.Items[0].Delta = -9000; state.Balances[0].NetOverride = 10000;
            Assert(ExcelReview.Compose(state, true).Contains("此项负偏差9000.00元") && ExcelReview.Compose(state, true).Contains("此项负偏差为1000.00元"), "手动切回偏差口径立即恢复原句式");
            var old = new ExcelReviewState { ModelVersion = 9, WordingMode = ExcelReview.NetWording }; ExcelReview.Normalize(old);
            Assert(old.ModelVersion == 11 && old.WordingMode == ExcelReview.DirectionWording && old.AmountDecimals == 2, "旧项目升级保持此前正负偏差及两位小数输出");

            string isolated = Path.Combine(output, "v0100_ui_data", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(isolated);
            var p = new Project { Values = new Dictionary<string, string>(example.Values), Excel = V9Example() };
            p.Values["报告日期"] = "2026年9月20日"; ExcelReview.StartLink(p, p.Excel);
            using (var form = new MainForm(true, isolated)) {
                V7Call(form, "LoadIntoForm", p); form.Show(); Application.DoEvents();
                var dates = V7Field<Dictionary<string, DateTimePicker>>(form, "dateInputs");
                Assert(dates.ContainsKey("报告日期") && dates["报告日期"].Checked && dates["报告日期"].Value.Date == new DateTime(2026, 9, 20), "报告日期与工期日期采用同一日历选择控件");
                dates["报告日期"].Value = new DateTime(2026, 10, 8); dates["报告日期"].Checked = true; V7Call(form, "Gather");
                Assert(p.Values["报告日期"] == "2026-10-08" && Core.Values(p, true)["报告日期"] == "2026年10月8日", "日历选择按项目保存并按模板中文日期输出");
                var inputs = V7Field<Dictionary<string, TextBox>>(form, "inputs");
                int projectHeight = inputs["项目名称"].Parent.Height, unitHeight = inputs["审核单位"].Parent.Height, dateHeight = dates["报告日期"].Parent.Height;
                Assert(Math.Abs(projectHeight - unitHeight) <= 4 && Math.Abs(unitHeight - dateHeight) <= 4, "项目信息单行项目采用一致的行高和间距（" + projectHeight + "/" + unitHeight + "/" + dateHeight + "）");
                var wording = V7Field<ComboBox>(form, "excelWording"); wording.SelectedIndex = 1; Application.DoEvents();
                Assert(p.Excel.WordingMode == ExcelReview.NetWording && Core.Get(p, ExcelReview.Targets[0]).Contains("净核减") && !Core.Get(p, ExcelReview.Targets[0]).Contains("负偏差"), "用户可在Excel页手动选择净核减增口径并立即联动正文");
                wording.SelectedIndex = 0; Assert(p.Excel.WordingMode == ExcelReview.DirectionWording && Core.Get(p, ExcelReview.Targets[0]).Contains("负偏差"), "用户可手动改回正负偏差口径");
                var tabs = V7Field<TabControl>(form, "tabs"); tabs.SelectedIndex = 4; Application.DoEvents();
                var split = V7Field<SplitContainer>(form, "excelSplit"); int before = split.Panel2.Height;
                int maxDistance = split.Height - split.Panel2MinSize - split.SplitterWidth;
                split.SplitterDistance = split.SplitterDistance + 50 <= maxDistance ? split.SplitterDistance + 50 : Math.Max(split.Panel1MinSize, split.SplitterDistance - 50); Application.DoEvents();
                Assert(split.Panel2.Height != before && V7Field<TabControl>(form, "excelPreviews").Visible, "Excel文字预览区域可拖动调整且小窗口不再强制隐藏");
                form.Screenshot(Path.Combine(output, "v0100_Excel预览可调及表述选择.png"), 4);
                tabs.SelectedIndex = 1; form.Show(); Application.DoEvents(); Control ancestor = dates["报告日期"];
                while (ancestor.Parent != null) { ancestor = ancestor.Parent; var scroll = ancestor as ScrollableControl; if (scroll != null && scroll.AutoScroll) { scroll.ScrollControlIntoView(dates["报告日期"]); break; } }
                Application.DoEvents(); using (var bmp = new Bitmap(form.Width, form.Height)) { form.DrawToBitmap(bmp, new Rectangle(0, 0, form.Width, form.Height)); bmp.Save(Path.Combine(output, "v0100_项目信息日期与间距.png")); } form.Hide();
                string chosen = Path.Combine(output, "用户选择的报告目录"); Directory.CreateDirectory(chosen);
                V7Call(form, "RememberExportDirectory", Path.Combine(chosen, "报告.docx"));
                Assert((string)V7Call(form, "LastExportDirectory") == chosen, "生成报告保存位置写入独立界面设置");
                form.Close();
            }
            using (var reopened = new MainForm(true, isolated)) Assert((string)V7Call(reopened, "LastExportDirectory") == Path.Combine(output, "用户选择的报告目录"), "重新打开软件继续使用上次报告保存目录");
            var saveProject = new Project { Values = new Dictionary<string, string>(example.Values), Excel = V9Example() }; saveProject.Excel.WordingMode = ExcelReview.NetWording;
            string save = Path.Combine(output, "v0100_文字口径保存.json"); Core.Save(save, saveProject);
            Assert(Core.Load(save).Excel.WordingMode == ExcelReview.NetWording, "手动选择的金额表述随项目保存重开");
        }
    }
}
