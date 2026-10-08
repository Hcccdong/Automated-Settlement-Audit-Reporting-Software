using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace SettlementReview {
    public static partial class Tests {
        static object V7Call(object obj, string method, params object[] args) { return obj.GetType().GetMethod(method, BindingFlags.Instance | BindingFlags.NonPublic).Invoke(obj, args); }
        static T V7Field<T>(object obj, string field) { return (T)obj.GetType().GetField(field, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(obj); }
        static void V7Mouse(DataGridView grid, string method, Point point) { typeof(Control).GetMethod(method, BindingFlags.Instance | BindingFlags.NonPublic).Invoke(grid, new object[] { new MouseEventArgs(MouseButtons.Left, 1, point.X, point.Y, 0) }); }
        static void CheckVersionSeven(string output, Project example, byte[] template) {
            var p = new Project { Values = new Dictionary<string, string>(example.Values) }; p.Values["有关事项说明"] = "";
            p.Values["商务标中标价"] = "900"; p.Values["招标控制价"] = "1100"; p.Values["其他费用"] = "100";
            Assert(Core.DiscountRate(p) == .2m && Core.DiscountText(Core.DiscountRate(p)) == "20.00%", "按指定扣除其他费用公式计算下浮率");
            p.Values["其他费用"] = "0"; Assert(Core.DiscountText(Core.DiscountRate(p)) == "18.18%", "其他费用为零且下浮率按百分数保留两位");
            p.Values["其他费用"] = "1100"; Reject(() => Core.DiscountRate(p), "下浮率分母为零明确提示");
            p.Values["其他费用"] = "100"; p.Values["商务标中标价"] = "1200"; Assert(Core.DiscountRate(p) < 0, "负下浮率不被静默截断成零"); p.Values["商务标中标价"] = "900";
            p.Values["下浮率模式"] = "自动";
            Assert(Core.Values(p, false)["工程下浮率"] == "20.00%", "正式导出采用最新自动计算值");
            p.Values["下浮率模式"] = "手动"; p.Values["工程下浮率"] = "19.50%（按合同约定）";
            Assert(Core.Values(p, false)["工程下浮率"] == "19.50%（按合同约定）", "手动下浮率及适用说明不被导出覆盖");
            p.Values["下浮率模式"] = "自动"; p.Values["其他费用"] = "1100"; Reject(() => Core.Export(p, template, false), "自动计算金额无效时不把旧下浮率输出为正式结果"); p.Values["其他费用"] = "100";
            using (var form = new MainForm(true)) {
                V7Call(form, "LoadIntoForm", p); form.Show(); Application.DoEvents();
                var inputs = V7Field<Dictionary<string, TextBox>>(form, "inputs"); var auto = V7Field<CheckBox>(form, "discountAuto");
                Assert(inputs["工程下浮率"].Text == "20.00%", "审核正文界面自动显示计算下浮率");
                inputs["工程下浮率"].Text = "17.25%";
                Assert(!auto.Checked && inputs["工程下浮率"].Text == "17.25%", "手动编辑结果立即切换模式且不被事件覆盖");
                inputs["商务标中标价"].Text = "800";
                Assert(inputs["工程下浮率"].Text == "17.25%", "手动模式修改计算输入仍保留手填结果");
                auto.Checked = true; Assert(inputs["工程下浮率"].Text == "30.00%", "恢复自动模式使用当前输入重新计算");
                form.Screenshot(Path.Combine(output, "v070_下浮率计算.png"), 3);
                inputs["工程下浮率"].Text = "22.50%"; V7Call(form, "Gather");
                string saved = Path.Combine(output, "v070_手动下浮率.json"); Core.Save(saved, p); V7Call(form, "LoadIntoForm", Core.Load(saved));
                Assert(!auto.Checked && inputs["工程下浮率"].Text == "22.50%" && inputs["商务标中标价"].Text == "800", "保存重开保留计算输入和手动模式");
                var state = ExcelReview.Parse(new List<ExcelSheet> {
                    ReviewTestSheet("甲楼_分部分项", "甲楼", "砌体", "200厚", "1000", "800"),
                    ReviewTestSheet("甲楼_技术措施", "甲楼", "模板", "木模板", "1000", "900"),
                    ReviewTestSheet("乙楼", "乙楼", "混凝土", "C30", "1000", "700"),
                    ReviewTestSheet("乙楼续表", "乙楼", "防水", "1.5厚", "1000", "950") }, "批量演示.xlsx");
                p.Excel = state; ExcelReview.StartLink(p, state); V7Call(form, "LoadIntoForm", p);
                var tabs = V7Field<TabControl>(form, "tabs"); tabs.SelectedIndex = 4; form.Show(); Application.DoEvents();
                var grid = V7Field<DataGridView>(form, "excelGrid");
                Assert(grid.MultiSelect && !grid.Columns.Cast<DataGridViewColumn>().Any(c => c.DataPropertyName == "Category"), "Excel支持多行选择且费用分类列已移除");
                Assert(!ExcelReview.Compose(state, true).Contains("分部分项：") && !ExcelReview.Compose(state, true).Contains("技术措施："), "待分类旧明细可生成且报告仅按子单位工程分组");
                var sheets = V7Field<ComboBox>(form, "excelSheets"); sheets.SelectedIndex = 1;
                V7Call(form, "SetVisibleInclude", 0); Assert(!state.Items[0].Include && state.Items.Skip(1).All(i => i.Include), "批量采用按钮只影响当前筛选明细");
                V7Call(form, "SetVisibleInclude", 2); Assert(state.Items.All(i => i.Include), "当前筛选反选恢复且不改变其他表");
                sheets.SelectedIndex = 0; V7Call(form, "SetVisibleInclude", 0); V7Call(form, "IncludeRange", 0, 2, true);
                Assert(state.Items.Take(3).All(i => i.Include) && !state.Items[3].Include, "Shift采用区间逻辑覆盖首尾且不影响区间外");
                V7Call(form, "SetVisibleInclude", 1);
                grid.ClearSelection(); grid.Rows[0].Selected = true; grid.Rows[2].Selected = true;
                V7Field<ComboBox>(form, "batchReason").SelectedItem = "单价偏差"; V7Call(form, "ApplySelectedReason");
                Assert(state.Items[0].Reason == "单价偏差" && state.Items[2].Reason == "单价偏差" && state.Items[1].Reason == "工程量偏差", "批量原因应用到选中行而非整表");
                int reasonCol = grid.Columns.Cast<DataGridViewColumn>().First(c => c.DataPropertyName == "Reason").Index;
                grid.CurrentCell = grid.Rows[0].Cells[reasonCol]; Application.DoEvents();
                Rectangle handle = (Rectangle)V7Call(form, "ReasonFillHandle");
                V7Mouse(grid, "OnMouseDown", new Point(handle.Left + 2, handle.Top + 2));
                grid.FirstDisplayedScrollingRowIndex = 3;
                Rectangle end = grid.GetCellDisplayRectangle(reasonCol, 3, true); Point endPoint = new Point(end.Left + 20, end.Top + 10);
                V7Mouse(grid, "OnMouseMove", endPoint);
                V7Mouse(grid, "OnMouseUp", endPoint);
                Assert(state.Items.All(i => i.Reason == "单价偏差"), "原因单元格右下角真实鼠标事件填充到目标行");
                state.Items[0].Reason = "工程量偏差"; grid.FirstDisplayedScrollingRowIndex = 0; grid.CurrentCell = grid.Rows[0].Cells[reasonCol];
                handle = (Rectangle)V7Call(form, "ReasonFillHandle"); V7Mouse(grid, "OnMouseDown", new Point(handle.Left + 2, handle.Top + 2));
                V7Call(form, "CancelReasonFill"); V7Mouse(grid, "OnMouseUp", endPoint);
                Assert(state.Items.Skip(1).All(i => i.Reason == "单价偏差"), "取消拖拽不会覆盖任何目标行的原因");
                state.Items[0].Reason = "待确认"; state.Items[1].Reason = ExcelReview.Reasons[3]; state.Items[1].Note = "";
                state.Items[2].Warning = "送审与审查名称不一致"; state.Items[2].Confirmed = false;
                V7Call(form, "ExcelChanged");
                Assert(ExcelReview.ProblemRows(state).Count == 3, "统一识别缺失原因其他说明和未核实数据三类问题");
                Assert((bool)V7Call(form, "LocateExcelProblems") && tabs.SelectedIndex == 4 && grid.Rows.Count == 3 && grid.Rows.Cast<DataGridViewRow>().All(r => r.ErrorText != ""), "定位切换Excel显示待处理行并保留逐行提示");
                tabs.SelectedIndex = 6; V7Call(form, "ExportReport", false);
                Assert(tabs.SelectedIndex == 4 && ReferenceEquals(grid.CurrentRow.DataBoundItem, state.Items[0]), "点击正式生成直接定位首个问题且无需确认勾选");
                form.Screenshot(Path.Combine(output, "v070_待处理定位.png"), 4);
                state.Items[0].Reason = "工程量偏差"; state.Items[1].Note = "依据已确认变更调整"; state.Items[2].Confirmed = true;
                V7Call(form, "ExcelChanged"); Assert(ExcelReview.ProblemRows(state).Count == 0 && grid.Rows.Cast<DataGridViewRow>().All(r => r.ErrorText == ""), "修复后颜色提示来源清除且校验同步通过");
                V7Field<CheckBox>(form, "excelProblemsOnly").Checked = false;
                grid.HorizontalScrollingOffset = 0; grid.CurrentCell = grid.Rows[0].Cells[reasonCol];
                form.Screenshot(Path.Combine(output, "v070_Excel批量操作.png"), 4);
                File.WriteAllBytes(Path.Combine(output, "v070_自动计算及单层分组.docx"), Core.Export(p, template, false));
                form.Screenshot(Path.Combine(output, "v070_报告导出.png"), 6);
                Assert(!ControlsBelow(tabs.TabPages[6].Controls).OfType<CheckBox>().Any(), "报告导出页无核对确认复选框");
                form.Close();
            }
        }
    }
}
