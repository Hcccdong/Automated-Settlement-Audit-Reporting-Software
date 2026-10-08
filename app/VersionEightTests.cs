using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;

namespace SettlementReview {
    public static partial class Tests {
        static Point V8Cell(DataGridView grid, int row, int column) { grid.FirstDisplayedScrollingRowIndex = row; var rect = grid.GetCellDisplayRectangle(column, row, true); return new Point(rect.Left + rect.Width / 2, rect.Top + rect.Height / 2); }
        static void CheckVersionEight(string output, Project example, byte[] template) {
            var p = new Project { Values = new Dictionary<string, string>(example.Values) }; p.Values["有关事项说明"] = "";
            p.Values["项目名称"] = "工程版本与勾选测试"; p.Values["报告编号"] = "温咨造〔2026〕008号";
            byte[] report = Core.Export(p, template, false); var doc = XDocument.Parse(Read(report, "word/document.xml"));
            var numberLine = doc.Descendants(Core.W + "p").Single(par => Core.RawParagraphText(par).Contains(p.Values["报告编号"]));
            Assert(Core.RawParagraphText(numberLine).Trim() == p.Values["报告编号"], "报告编号行不再自动拼接工程名称或分隔符");
            var original = XDocument.Parse(Read(template, "word/document.xml")).Descendants(Core.W + "p").First(par => Core.RawParagraphText(par).Contains("{{报告编号}}"));
            Assert(XNode.DeepEquals(original.Element(Core.W + "pPr"), numberLine.Element(Core.W + "pPr")), "去编号后缀保留原段落边线与格式");
            Assert(doc.Descendants(Core.W + "p").Any(par => Core.RawParagraphText(par).Contains("关于" + p.Values["项目名称"])), "报告标题中的工程名称仍然保留");
            p.Values["报告编号"] = "自定008号-补充后缀";
            Assert(Core.Preview(p, template).Contains("自定008号-补充后缀") && !Core.Preview(p, template).Contains("自定008号-补充后缀-工程版本"), "手工编号后缀完整输出且不叠加名称");
            File.WriteAllBytes(Path.Combine(output, "v080_编号及格式验证.docx"), Core.Export(p, template, false));
            var old = new Project { Values = new Dictionary<string, string>(p.Values), Updated = "2026-09-15 10:00" }; p.Updated = "2026-09-16 10:00";
            var other = new Project(); other.Values["项目名称"] = "另一个工程";
            var groups = Core.GroupProjects(new[] { old, p, other, new Project(), new Project() }, p.Id);
            Assert(groups.Count == 4 && groups.First().Versions.Count == 2 && groups.First().Latest.Id == p.Id, "旧同名记录合并为一个工程入口且空白工程不误合并");
            string dataRoot = Path.Combine(output, "isolated_data", Guid.NewGuid().ToString("N"));
            using (var form = new MainForm(true, dataRoot)) {
                V7Call(form, "LoadIntoForm", p); V7Call(form, "CopyProject");
                string initialId = p.Id; string[] snapshots = Directory.GetFiles(Path.Combine(dataRoot, "copies", p.Id), "*.json");
                var projects = V7Field<ListBox>(form, "projects");
                Assert(p.Id == initialId && Directory.GetFiles(Path.Combine(dataRoot, "projects"), "*.json").Length == 1 && projects.Items.Count == 1 && snapshots.Length == 1, "保存审核副本保留工程ID且左侧只有一个名字");
                var inputs = V7Field<Dictionary<string, TextBox>>(form, "inputs"); inputs["报告编号"].Text = "后续工作编号"; V7Call(form, "SaveProject");
                Assert(Core.Get(Core.Load(snapshots[0]), "报告编号") == "自定008号-补充后缀", "继续保存工作项目不会覆盖历史副本");
                Core.Save(Path.Combine(dataRoot, "projects", old.Id + ".json"), old); V7Call(form, "RefreshProjects");
                Assert(projects.Items.Count == 1 && ((ProjectGroup)projects.Items[0]).Versions.Count == 2, "已有旧版同名副本在界面合并且原文件保留");
                var versions = (List<ProjectVersion>)V7Call(form, "ProjectVersionFiles", projects.Items[0]);
                Assert(versions.Count == 3 && versions.Any(v => v.Path == snapshots[0]), "版本入口包含旧同名记录和新审核副本");
                using (var dialog = (Form)V7Call(form, "CreateProjectVersionDialog", projects.Items[0])) {
                    dialog.Show(); Application.DoEvents();
                    using (var bitmap = new Bitmap(dialog.Width, dialog.Height)) { dialog.DrawToBitmap(bitmap, new Rectangle(0, 0, dialog.Width, dialog.Height)); bitmap.Save(Path.Combine(output, "v080_审核版本窗口.png")); } dialog.Close();
                }
                V7Call(form, "OpenProjectVersion", versions.First(v => v.Path == snapshots[0]));
                Assert(inputs["报告编号"].Text == "自定008号-补充后缀", "可重新打开历史副本的数据继续审核");
                var state = ExcelReview.Parse(new List<ExcelSheet> {
                    ReviewTestSheet("明细甲", "1号楼", "砌体", "200厚", "1000", "800"),
                    ReviewTestSheet("明细乙", "1号楼", "模板", "木模板", "1000", "900"),
                    ReviewTestSheet("明细丙", "2号楼", "混凝土", "C30", "1000", "700"),
                    ReviewTestSheet("明细丁", "2号楼", "防水", "1.5厚", "1000", "950") }, "多行勾选.xlsx");
                p.Excel = state; ExcelReview.StartLink(p, state); V7Call(form, "LoadIntoForm", p);
                var tabs = V7Field<TabControl>(form, "tabs"); tabs.SelectedIndex = 4; form.Show(); Application.DoEvents();
                var grid = V7Field<DataGridView>(form, "excelGrid"); V7Call(form, "SetVisibleInclude", 0);
                grid.ClearSelection(); grid.Rows[0].Selected = true; grid.Rows[1].Selected = true;
                Point click = V8Cell(grid, 1, 0); V7Mouse(grid, "OnMouseDown", click); V7Mouse(grid, "OnMouseUp", click);
                Assert(state.Items.Take(2).All(i => i.Include) && state.Items.Skip(2).All(i => !i.Include) && grid.SelectedRows.Count == 2, "多行选中后鼠标点击采用格批量勾选且选择不收缩");
                V7Mouse(grid, "OnMouseDown", click); V7Mouse(grid, "OnMouseUp", click);
                Assert(state.Items.All(i => !i.Include), "再次点击选中行采用格批量取消而非仅取消单行");
                grid.ClearSelection(); grid.Rows[0].Selected = true; grid.Rows[2].Selected = true;
                click = V8Cell(grid, 0, 0); V7Mouse(grid, "OnMouseDown", click); V7Mouse(grid, "OnMouseUp", click);
                Assert(state.Items[0].Include && state.Items[2].Include && !state.Items[1].Include && !state.Items[3].Include, "不连续选中行也可批量勾选且不波及中间行");
                click = V8Cell(grid, 3, 0); V7Mouse(grid, "OnMouseDown", click); V7Mouse(grid, "OnMouseUp", click);
                Assert(state.Items[0].Include && state.Items[2].Include && state.Items[3].Include && !state.Items[1].Include && grid.SelectedRows.Count == 1, "点击未选中的行只处理新行而非之前的选择集合");
                V7Call(form, "HandleIncludeClick", V8Cell(grid, 1, 0), true);
                Assert(state.Items.All(i => i.Include), "Shift采用区间继续支持首尾统一勾选");
                Assert(grid.Columns.Cast<DataGridViewColumn>().Single(c => c.DataPropertyName == "Group").HeaderText == "单项工程（可批量覆盖）" && ExcelReview.RowIssue(new ReviewItem()).Contains("单项工程"), "表头及缺失提示统一改为单项工程");
                V7Call(form, "SaveProject"); grid.FirstDisplayedScrollingRowIndex = 0; grid.ClearSelection(); grid.Rows[0].Selected = true; grid.Rows[1].Selected = true;
                form.Screenshot(Path.Combine(output, "v080_工程列表与多行采用.png"), 4);
                form.Screenshot(Path.Combine(output, "v080_编号预览.png"), 6);
                form.Close();
            }
        }
    }
}
