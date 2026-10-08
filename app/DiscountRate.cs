using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace SettlementReview {
    public static partial class Core {
        public static readonly string[] DiscountInputs = { "商务标中标价", "招标控制价", "其他费用" };
        public static decimal DiscountRate(Project p) {
            decimal bid = Money(Get(p, DiscountInputs[0]), DiscountInputs[0]);
            decimal limit = Money(Get(p, DiscountInputs[1]), DiscountInputs[1]);
            decimal other = Money(Get(p, DiscountInputs[2]), DiscountInputs[2]);
            if (limit == other) throw new InvalidDataException("招标控制价减其他费用为零，无法计算下浮率。");
            return 1m - (bid - other) / (limit - other);
        }
        public static string DiscountText(decimal rate) { return (decimal.Round(rate * 100m, 2, MidpointRounding.AwayFromZero)).ToString("0.00", CultureInfo.InvariantCulture) + "%"; }
    }
    public partial class MainForm {
        readonly CheckBox discountAuto = new CheckBox();
        readonly Label discountStatus = new Label();
        bool discountUpdating;
        void BuildDiscountInputs(Panel cell, TextBox result) {
            cell.Height += 206;
            var calculator = new Panel { Dock = DockStyle.Top, Height = 206 };
            var grid = new TableLayoutPanel { Dock = DockStyle.Top, Height = 74, ColumnCount = 3, RowCount = 2 };
            for (int i = 0; i < 3; i++) {
                grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F / 3));
                grid.Controls.Add(new Label { Text = Core.DiscountInputs[i] + "（元）", Dock = DockStyle.Fill, AutoSize = true }, i, 0);
                var input = new TextBox { Dock = DockStyle.Top, Margin = new Padding(0, 5, 14, 0) };
                inputs[Core.DiscountInputs[i]] = input; input.TextChanged += (s,e) => Changed(); grid.Controls.Add(input, i, 1);
            }
            discountAuto.Text = "自动更新下浮率（直接修改下方结果即转为手动）"; discountAuto.Dock = DockStyle.Top; discountAuto.Height = 34;
            discountAuto.CheckedChanged += (s,e) => { if (!loading && !discountUpdating) Changed(); };
            discountStatus.Dock = DockStyle.Fill; discountStatus.Padding = new Padding(0, 4, 8, 0);
            calculator.Controls.Add(discountStatus); calculator.Controls.Add(discountAuto); calculator.Controls.Add(grid);
            var footer = new Panel { Dock = DockStyle.Bottom, Height = 94 };
            footer.Controls.Add(new Label { Text = "报告采用的下浮率（可手动修改，也可补充适用说明）", Dock = DockStyle.Bottom, Height = 38, ForeColor = Color.DimGray });
            cell.Controls.Remove(result); footer.Controls.Add(result); cell.Controls.Clear(); cell.Controls.Add(footer); cell.Controls.Add(calculator);
        }
        void UpdateDiscount() {
            if (discountUpdating || !inputs.ContainsKey("工程下浮率")) return;
            discountUpdating = true;
            try {
                bool complete = true; foreach (string key in Core.DiscountInputs) if (Core.Get(current, key) == "") complete = false;
                if (!complete) { discountStatus.ForeColor = Color.DimGray; discountStatus.Text = "计算式：1 −（商务标中标价 − 其他费用）÷（招标控制价 − 其他费用）\n填写三项金额后计算；其他费用没有则填0。下方可直接填写报告采用的下浮率。"; return; }
                decimal rate = Core.DiscountRate(current); string computed = Core.DiscountText(rate);
                discountStatus.ForeColor = rate < 0 || rate > 1 ? Color.Firebrick : Color.FromArgb(27, 78, 103);
                discountStatus.Text = "公式计算结果：" + computed + (discountAuto.Checked ? " · 已自动填入下方" : " · 当前采用手动值，修改金额不会覆盖") + "\n1 −（商务标中标价 − 其他费用）÷（招标控制价 − 其他费用）" + (rate < 0 || rate > 1 ? "\n结果超出0%～100%，请核对金额；未强行改为零。" : "");
                if (discountAuto.Checked) { bool wasLoading = loading; loading = true; try { inputs["工程下浮率"].Text = computed; current.Values["工程下浮率"] = computed; } finally { loading = wasLoading; } }
            } catch (Exception ex) { discountStatus.ForeColor = Color.Firebrick; discountStatus.Text = ex.Message + "\n未更新下方结果；可更正金额或直接手动填写。"; }
            finally { discountUpdating = false; }
        }
    }
}
