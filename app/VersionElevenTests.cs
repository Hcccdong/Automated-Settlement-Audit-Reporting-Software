using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SettlementReview {
    public static partial class Tests {
        static void CheckVersionEleven(string output, Project example, byte[] template) {
            var automatic = new Project { Values = new Dictionary<string, string>(example.Values) };
            automatic.Values["送审金额"] = "1000"; automatic.Values["审定金额"] = "1020"; automatic.Values["净核减金额"] = "";
            var values = Core.Values(automatic, false);
            Assert(values["净核减金额"] == "-20.00" && values["净额方向"] == "净核增" && values["净额绝对金额"] == "20.00" && values["审核结论"].Contains("净核增金额为20.00元"), "金额结论自动计算并按符号写净核增");
            automatic.Values["净核减金额"] = "-25"; values = Core.Values(automatic, false);
            Assert(values["审核结论"].Contains("净核增金额为25.00元") && Core.AmountCheck(automatic).Contains("提示："), "人工净额覆盖自动值并保留验算提示");

            var state = V9Example(); state.WordingMode = ExcelReview.NetWording; state.AmountDecimals = 0; state.Threshold = 0;
            state.Items[0].Delta = -9000.5m; state.Items[0].AuditAmount = state.Items[0].SentAmount - 9000.5m;
            state.Items[1].Delta = -100.5m; state.Items[1].AuditAmount = state.Items[1].SentAmount - 100.5m;
            state.Balances[0].NetOverride = 10000m;
            string zero = ExcelReview.Compose(state, true);
            Assert(zero.Contains("净核减9001元") && zero.Contains("净核减101元") && zero.Contains("净核减898元"), "零位小数逐条四舍五入并以末条配平一万元");
            var balance = ExcelReview.CalculateBalance(state, state.Balances[0]);
            Assert(balance.Listed == -9102m && balance.Residual == -898m && balance.Listed + balance.Residual == balance.TotalDelta, "金额配平采用报告显示精度计算");
            state.AmountDecimals = 1; string one = ExcelReview.Compose(state, true);
            Assert(one.Contains("送审金额为15000.0元") && one.Contains("净核减9000.5元") && one.Contains("净核减899.0元"), "一位小数同步作用于送审审核金额、偏差和末条");
            Assert(ExcelReview.DirectionText(state, state.Items[0]) == "净核减", "净额表述时方向列显示净核减");
            state.WordingMode = ExcelReview.DirectionWording;
            Assert(ExcelReview.DirectionText(state, state.Items[0]) == "负偏差", "正负偏差表述时方向列同步显示负偏差");
            state.AmountDecimals = 4; Assert(ExcelReview.Money(state, 1.23455m) == "1.2346", "四位小数使用四舍五入而非截断");

            string isolated = Path.Combine(output, "v0110_ui_data", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(isolated);
            var p = new Project { Values = new Dictionary<string, string>(example.Values), Excel = state };
            p.Values["净额模式"] = "自动"; p.Values["送审金额"] = "1000"; p.Values["审定金额"] = "1020"; p.Values["净核减金额"] = "";
            ExcelReview.StartLink(p, state);
            using (var form = new MainForm(true, isolated)) {
                V7Call(form, "LoadIntoForm", p); form.Show(); Application.DoEvents();
                var inputs = V7Field<Dictionary<string, TextBox>>(form, "inputs"); var auto = V7Field<CheckBox>(form, "netAmountAuto");
                Assert(auto.Checked && inputs["净核减金额"].Text == "-20.00" && p.Values["净额模式"] == "自动", "金额页打开后自动生成净核增数值");
                inputs["净核减金额"].Text = "-25.00"; Application.DoEvents();
                Assert(!auto.Checked && p.Values["净核减金额"] == "-25.00" && p.Values["净额模式"] == "手动", "直接修改净额自动切换为手动并保留输入");
                inputs["送审金额"].Text = "1100"; Application.DoEvents(); Assert(inputs["净核减金额"].Text == "-25.00", "手动净额不被送审审核金额变化覆盖");
                auto.Checked = true; Application.DoEvents(); Assert(inputs["净核减金额"].Text == "80.00", "重新勾选自动后立即恢复计算值");

                var wording = V7Field<ComboBox>(form, "excelWording"); var decimals = V7Field<ComboBox>(form, "excelDecimals"); var grid = V7Field<DataGridView>(form, "excelGrid");
                wording.SelectedIndex = 1; decimals.SelectedIndex = 0; Application.DoEvents();
                var directionColumn = grid.Columns.Cast<DataGridViewColumn>().Single(c => c.DataPropertyName == "Direction");
                var deltaColumn = grid.Columns.Cast<DataGridViewColumn>().Single(c => c.DataPropertyName == "Delta");
                Assert(state.AmountDecimals == 0 && deltaColumn.DefaultCellStyle.Format == "F0" && Convert.ToString(grid.Rows[0].Cells[directionColumn.Index].FormattedValue) == "净核减", "界面小数位和偏差方向随金额表述即时更新");
                Assert(ControlsBelow(form.Controls).OfType<Button>().Any(b => b.Text == "金额配平") && !ControlsBelow(form.Controls).OfType<Button>().Any(b => b.Text == "工程总额与差额"), "功能入口已更名为金额配平");
                using (var dialog = (Form)V7Call(form, "CreateBalanceDialog")) {
                    Assert(dialog.Text == "金额配平", "配平窗口标题同步更名");
                    var balanceGrid = ControlsBelow(dialog.Controls).OfType<DataGridView>().Single();
                    Assert(Convert.ToString(balanceGrid.Rows[0].Cells["Total"].Value) == "10000" && Convert.ToString(balanceGrid.Rows[0].Cells["Residual"].Value) == "898", "配平窗口同步显示所选小数位及四舍五入差额");
                }
                form.Screenshot(Path.Combine(output, "v0110_金额方向小数位与配平.png"), 4);
                form.Screenshot(Path.Combine(output, "v0110_净额自动与手动.png"), 5);
                form.Close();
            }
            string saved = Path.Combine(output, "v0110_设置保存.json"); Core.Save(saved, p); var reopened = Core.Load(saved);
            Assert(reopened.Excel.AmountDecimals == 0 && reopened.Excel.WordingMode == ExcelReview.NetWording && Core.Get(reopened, "净额模式") == "自动", "小数位、金额表述和净额模式随项目保存");
        }
    }
}
