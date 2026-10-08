using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SettlementReview {
    public static partial class Tests {
        static void CheckVersionFourteen(string output) {
            var a = new ReviewItem { Group = "原工程", Sheet = "明细", CalculationKey = "key", Include = true, Delta = -10m, Name = "甲", Warning = "", Reason = ExcelReview.Reasons[0], Target = ExcelReview.Targets[0], Category = ExcelReview.Categories[0] };
            var hidden = new ReviewItem { Group = "原工程", Sheet = "明细", CalculationKey = "key", Include = false, Delta = 0m, Name = "隐藏零差额", Warning = "", Reason = ExcelReview.Reasons[0], Target = ExcelReview.Targets[0], Category = ExcelReview.Categories[0] };
            var state = new ExcelReviewState { ModelVersion = 11, AmountDecimals = 2, Items = new List<ReviewItem> { a, hidden }, Totals = new List<ReviewTotal> { new ReviewTotal { Group = "原工程", Key = "key", Delta = -15m, Source = "计算表", Message = "", DetailSheets = new List<string> { "明细" } } }, Balances = new List<ReviewBalance> { new ReviewBalance { Group = "原工程", Enabled = true, Target = ExcelReview.Targets[0], Reason = ExcelReview.FeeReason } } };
            ReviewGroupEditor.RenameGroup(state, new[] { a }, "新工程");
            Assert(hidden.Group == "新工程" && state.Totals[0].Group == "新工程" && state.Balances[0].Group == "新工程", "筛选下改名包括隐藏同组行并同步计算表补差");
            var result = ExcelReview.CalculateBalance(state, state.Balances[0]);
            Assert(result.Issue == "" && result.TotalDelta == -15m && result.Residual == -5m, "筛选改名后配平继续生效且金额不变");
            state.Totals[0].Group = "旧版错误名称"; state.Balances[0].Group = "旧版错误名称"; ExcelReview.Normalize(state);
            Assert(state.Totals[0].Group == "新工程" && state.Balances[0].Group == "新工程", "旧项目按稳定计算表key修复名称归属");
            string file = Path.Combine(output, "v014_改名配平.json"); var project = new Project { Excel = state }; Core.Save(file, project);
            var reopened = Core.Load(file); ExcelReview.Normalize(reopened.Excel);
            Assert(ExcelReview.CalculateBalance(reopened.Excel, reopened.Excel.Balances[0]).Residual == -5m, "改名配平保存重开金额一致");
            string library = Path.Combine(output, "v014_basis_templates.json"); var templates = BasisTemplateStore.Load(library);
            BasisTemplateStore.Save(library, templates, "template-a", "自建依据模板", "合同依据第一条\r\n合同依据第二条");
            BasisTemplateStore.Save(library, templates, "template-a", "改名后的依据模板", "修改后的模板内容");
            Assert(templates.Count == 1 && BasisTemplateStore.Load(library)[0].Name == "改名后的依据模板", "自建依据模板改名按ID更新不会重复新增");
            Assert(BasisTemplateStore.Load(library)[0].Text == "修改后的模板内容", "依据模板内容修改后持久保存");
            Reject(() => BasisTemplateStore.Save(library, templates, "template-b", "改名后的依据模板", "其他内容"), "模板名称重复拒绝覆盖其他模板");
            Assert(templates.Count == 1 && BasisTemplateStore.Load(library)[0].Text == "修改后的模板内容", "模板保存失败保留既有内容");
            string isolated = Path.Combine(output, "v014_ui_data"); Directory.CreateDirectory(isolated);
            using (var form = new MainForm(true, isolated)) {
                V7Call(form, "LoadIntoForm", project); form.Show(); var tabs = V7Field<TabControl>(form, "tabs"); tabs.SelectedIndex = 4; Application.DoEvents();
                var grid = V7Field<DataGridView>(form, "excelGrid");
                grid.CurrentCell = grid.Rows[0].Cells[grid.Columns.Cast<DataGridViewColumn>().Single(c => c.DataPropertyName == "Group").Index]; grid.BeginEdit(true);
                ((TextBox)grid.EditingControl).Text = "直接编辑工程"; grid.EndEdit(); Application.DoEvents();
                Assert(state.Items.All(i => i.Group == "直接编辑工程") && state.Totals[0].Group == "直接编辑工程" && ExcelReview.CalculateBalance(state, state.Balances[0]).Issue == "", "直接编辑名称也同步完整工程与配平");
                var bar = V7Field<FlowLayoutPanel>(form, "excelSelectionBar"); var select = bar.Controls.OfType<Button>().First(b => b.Text == "全选");
                Assert(select.Font.Size < form.Font.Size && select.PointToScreen(Point.Empty).Y < grid.PointToScreen(Point.Empty).Y, "全选按钮移到表格项目名称上方并缩小字体");
                form.Screenshot(Path.Combine(output, "v014_全选与配平修复.png"), 4);
                var report = V7Field<Dictionary<string, TextBox>>(form, "inputs")["审核依据及原则"];
                using (var dialog = (Form)V7Call(form, "CreateBasisTemplateDialog", report)) {
                    dialog.Show(form); Application.DoEvents();
                    var controls = ControlsBelow(dialog.Controls).ToList(); var name = controls.OfType<TextBox>().Single(t => t.Name == "BasisTemplateName"); var text = controls.OfType<TextBox>().Single(t => t.Name == "BasisTemplateText");
                    name.Text = "用户新增模板"; text.Text = "用户新增条目"; controls.OfType<Button>().Single(b => b.Text == "保存名称与内容").PerformClick();
                    string personalFile = Path.Combine(isolated, "basis-templates.json"); Assert(BasisTemplateStore.Load(personalFile).Single().Name == "用户新增模板", "管理窗口新增模板保存到隔离模板库");
                    name.Text = "用户改名模板"; controls.OfType<Button>().Single(b => b.Text == "保存名称与内容").PerformClick(); Assert(BasisTemplateStore.Load(personalFile).Count == 1 && BasisTemplateStore.Load(personalFile)[0].Name == "用户改名模板", "管理窗口改名保存保持同一模板");
                    using (var bitmap = new Bitmap(dialog.Width, dialog.Height)) { dialog.DrawToBitmap(bitmap, new Rectangle(0,0,bitmap.Width,bitmap.Height)); bitmap.Save(Path.Combine(output, "v014_个人依据模板管理.png")); }
                    controls.OfType<Button>().Single(b => b.Text == "使用当前内容").PerformClick(); Assert(report.Text == "用户新增条目", "个人依据模板可应用到当前报告");
                }
            }
        }
    }
}
