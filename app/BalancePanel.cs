using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace SettlementReview {
    public partial class MainForm {
        readonly CheckBox excelConcise = new CheckBox();
        void ShowBalanceDialog() {
            if (current.Excel == null) return;
            excelGrid.EndEdit(); ExcelReview.Normalize(current.Excel);
            using (var dialog = CreateBalanceDialog()) dialog.ShowDialog(this);
            Changed(); UpdateExcelPreview();
        }
        Form CreateBalanceDialog() {
            var state = current.Excel;
            foreach (var group in state.Items.Select(i => i.Group).GroupBy(ExcelReview.Norm)) if (!state.Balances.Any(b => ExcelReview.Norm(b.Group) == group.Key)) state.Balances.Add(new ReviewBalance { Group = group.First(), Target = ExcelReview.Targets[0], Reason = ExcelReview.FeeReason });
            var dialog = new Form { Text = "金额配平", Font = Font, Size = new Size(1200, 710), StartPosition = FormStartPosition.CenterParent, MinimumSize = new Size(900, 560), AutoScaleMode = AutoScaleMode.None, BackColor = Color.FromArgb(245, 247, 250) };
            var grid = new DataGridView { Dock = DockStyle.Fill, AutoGenerateColumns = false, AllowUserToAddRows = false, AllowUserToDeleteRows = false, RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect, MultiSelect = false, BackgroundColor = Color.White, RowTemplate = { Height = 34 }, ColumnHeadersHeight = 42 };
            grid.EnableHeadersVisualStyles = false; grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(232, 240, 247); grid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(31, 56, 78); grid.ColumnHeadersDefaultCellStyle.Font = new Font(Font, FontStyle.Bold);
            grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(218, 235, 249); grid.DefaultCellStyle.SelectionForeColor = Color.FromArgb(24, 49, 70); grid.GridColor = Color.FromArgb(231, 236, 242);
            grid.Columns.Add(new DataGridViewCheckBoxColumn { Name = "Enabled", HeaderText = "末条补差", Width = 85, Frozen = true });
            string[] names = { "Group", "Total", "Listed", "Residual", "After", "Check", "Override" };
            string[] labels = { "单项工程", "总额净核减", "已列净核减", "末条净核减", "报告净核减", "核对结果", "人工总额（选填）" };
            for (int i = 0; i < names.Length; i++) grid.Columns.Add(new DataGridViewTextBoxColumn { Name = names[i], HeaderText = labels[i], ReadOnly = names[i] != "Override", Width = i == 0 ? 160 : i == 5 ? 175 : 125, SortMode = DataGridViewColumnSortMode.NotSortable });
            grid.Columns["Group"].Frozen = true;
            foreach (string name in new[] { "Total", "Listed", "Residual", "After", "Override" }) grid.Columns[name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            grid.Columns.Add(new DataGridViewComboBoxColumn { Name = "Target", HeaderText = "末条报告归属", Width = 190, DataSource = ExcelReview.Targets });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Reason", HeaderText = "末条说明（可修改）", Width = 540 });
            var details = new TextBox { Dock = DockStyle.Bottom, Height = 195, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, BackColor = Color.FromArgb(245, 249, 252) };
            var banner = new Label { Dock = DockStyle.Top, Height = 70, Padding = new Padding(12, 8, 12, 0), Text = "金额单位：元；净核减正数表示核减，负数表示核增。末条差额 = 计算表总额 − 已列报告明细（按净核减口径）。\n门槛和勾选项变化后自动重算。同一工程跨多个报告归属时仅补一次差额；向右滚动可修改归属及末条说明。" };
            bool busy = true;
            Action<DataGridViewRow> refresh = row => {
                var b = (ReviewBalance)row.Tag; var r = ExcelReview.CalculateBalance(state, b);
                row.Cells["Total"].Value = r.TotalDelta.HasValue ? ExcelReview.Money(state, -r.TotalDelta.Value) : "未识别";
                row.Cells["Listed"].Value = ExcelReview.Money(state, -r.Listed); row.Cells["Residual"].Value = r.TotalDelta.HasValue ? ExcelReview.Money(state, -r.Residual) : "未识别";
                row.Cells["After"].Value = ExcelReview.Money(state, -(r.Listed + (b.Enabled && r.Issue == "" ? r.Residual : 0)));
                row.Cells["Check"].Value = !b.Enabled ? "未启用末条补差" : r.Issue != "" ? r.Issue : "金额一致";
                row.DefaultCellStyle.BackColor = b.Enabled && r.Issue != "" ? Color.MistyRose : Color.White;
                row.DefaultCellStyle.SelectionBackColor = b.Enabled && r.Issue != "" ? Color.LightSalmon : Color.FromArgb(218, 235, 249);
                row.Cells["Check"].ToolTipText = !b.Enabled ? "当前未启用末条补差。" : r.Issue == "" ? "报告净核减与单项工程总额相同，差额为" + ExcelReview.Money(state, 0) + "元。" : r.Issue;
            };
            Action showSource = () => {
                if (grid.CurrentRow == null || grid.CurrentRow.Tag == null) return;
                var b = (ReviewBalance)grid.CurrentRow.Tag; var r = ExcelReview.CalculateBalance(state, b);
                details.Text = b.Group + "\r\n" + r.Source + "\r\n未列入的明细净核减（含低于门槛和取消采用）：" + ExcelReview.Money(state, -r.Omitted) + " 元\r\n计算表与全部已读取明细的净差额：" + (r.TotalDelta.HasValue ? ExcelReview.Money(state, -r.NonDetail) : "未识别") + " 元\r\n以上两部分合计为末条净核减；统一使用所填末条说明。\r\n" + (r.Issue == "" ? "" : "待处理：" + r.Issue) + "\r\n末条说明：" + b.Reason;
            };
            foreach (var balance in state.Balances) {
                int index = grid.Rows.Add(balance.Enabled, balance.Group, "", "", "", "", "", balance.NetOverride.HasValue ? ExcelReview.Money(state, balance.NetOverride.Value) : "", balance.Target, balance.Reason ?? ExcelReview.FeeReason);
                grid.Rows[index].Tag = balance; refresh(grid.Rows[index]);
            }
            busy = false;
            grid.CurrentCellDirtyStateChanged += (s,e) => { if (grid.IsCurrentCellDirty) grid.CommitEdit(DataGridViewDataErrorContexts.Commit); };
            grid.CellValidating += (s,e) => {
                if (e.RowIndex < 0 || grid.Columns[e.ColumnIndex].Name != "Override") return;
                string text = Convert.ToString(e.FormattedValue).Trim();
                try { if (text != "") Core.SignedMoney(text); grid.Rows[e.RowIndex].ErrorText = ""; }
                catch { e.Cancel = true; grid.Rows[e.RowIndex].ErrorText = "请填写最多两位小数的净核减金额，负数为净核增；留空使用计算表总额。"; }
            };
            grid.CellValueChanged += (s,e) => {
                if (busy || e.RowIndex < 0 || grid.Rows[e.RowIndex].Tag == null) return;
                var row = grid.Rows[e.RowIndex]; var b = (ReviewBalance)row.Tag;
                string key = grid.Columns[e.ColumnIndex].Name;
                if (key == "Enabled") b.Enabled = Convert.ToBoolean(row.Cells[key].Value);
                else if (key == "Target") b.Target = Convert.ToString(row.Cells[key].Value);
                else if (key == "Reason") b.Reason = Convert.ToString(row.Cells[key].Value);
                else if (key == "Override") { string text = Convert.ToString(row.Cells[key].Value).Trim(); try { b.NetOverride = text == "" ? (decimal?)null : Core.SignedMoney(text); } catch { row.ErrorText = "请填写最多两位小数的净核减金额，负数为净核增。"; return; } }
                else return;
                busy = true; refresh(row); busy = false; showSource(); Changed(); UpdateExcelPreview();
            };
            grid.SelectionChanged += (s,e) => showSource();
            grid.DataError += (s,e) => { e.ThrowException = false; if (e.RowIndex >= 0) grid.Rows[e.RowIndex].ErrorText = "请检查填写的总额或报告归属。"; };
            dialog.FormClosing += (s,e) => { if (!grid.EndEdit()) e.Cancel = true; };
            var footer = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 52, Padding = new Padding(8) };
            footer.Controls.Add(Button("完成", (s,e) => dialog.Close(), 100));
            footer.Controls.Add(new Label { AutoSize = true, Margin = new Padding(10, 8, 0, 0), Text = "留空人工总额即恢复计算表值；勾选末条补差后才写入报告。" });
            dialog.Controls.Add(grid); dialog.Controls.Add(details); dialog.Controls.Add(footer); dialog.Controls.Add(banner);
            dialog.Load += (s,e) => {
                float factor = dialog.DeviceDpi / 96F; if (factor != 1F) dialog.Scale(new SizeF(factor, factor));
                foreach (DataGridViewColumn col in grid.Columns) col.Width = (int)(col.Width * factor);
                grid.ColumnHeadersHeight = (int)(42 * factor); foreach (DataGridViewRow row in grid.Rows) row.Height = (int)(34 * factor);
                dialog.AutoScaleMode = AutoScaleMode.Dpi; dialog.AutoScaleDimensions = new SizeF(dialog.DeviceDpi, dialog.DeviceDpi);
                var area = Screen.FromControl(dialog).WorkingArea; dialog.MinimumSize = new Size(Math.Min(dialog.MinimumSize.Width, area.Width), Math.Min(dialog.MinimumSize.Height, area.Height)); dialog.Size = new Size(Math.Min(dialog.Width, area.Width), Math.Min(dialog.Height, area.Height));
            };
            showSource(); return dialog;
        }
    }
}
