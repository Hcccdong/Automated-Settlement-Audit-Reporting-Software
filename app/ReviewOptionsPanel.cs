using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace SettlementReview {
    public partial class MainForm {
        readonly ComboBox batchGroup = new ComboBox();
        void BuildGroupControls(FlowLayoutPanel options) {
            options.SetFlowBreak(options.Controls[options.Controls.Count - 1], true);
            options.Controls.Add(new Label { Text = "单项工程名称", AutoSize = true, Margin = new Padding(0, 8, 6, 0) });
            batchGroup.DropDownStyle = ComboBoxStyle.DropDown; batchGroup.Width = 220;
            batchGroup.DropDown += (s,e) => RefreshGroupChoices(); options.Controls.Add(batchGroup);
            options.Controls.Add(Button("覆盖选中行名称", (s,e) => ApplySelectedGroup(), 145));
            var menu = new ContextMenuStrip();
            menu.Opening += (s,e) => {
                menu.Items.Clear();
                if (current.Excel == null) { e.Cancel = true; return; }
                foreach (string group in current.Excel.Items.Select(i => i.Group).Concat(current.Excel.Totals.Select(t => t.Group)).Where(g => !string.IsNullOrWhiteSpace(g)).Distinct()) {
                    string choice = group; menu.Items.Add(choice, null, (sender,args) => { batchGroup.Text = choice; ApplySelectedGroup(); });
                }
                if (menu.Items.Count == 0) e.Cancel = true;
            };
            excelGrid.ContextMenuStrip = menu;
            excelGrid.CellMouseDown += (s,e) => { if (e.Button == MouseButtons.Right && e.RowIndex >= 0 && !excelGrid.Rows[e.RowIndex].Selected) { excelGrid.ClearSelection(); excelGrid.Rows[e.RowIndex].Selected = true; } };
        }
        void RefreshGroupChoices() {
            if (current.Excel == null) return;
            string text = batchGroup.Text; batchGroup.Items.Clear();
            batchGroup.Items.AddRange(current.Excel.Items.Select(i => i.Group).Concat(current.Excel.Totals.Select(t => t.Group)).Where(g => !string.IsNullOrWhiteSpace(g)).Distinct().Cast<object>().ToArray()); batchGroup.Text = text;
        }
        void ApplySelectedGroup() { Run(() => {
            var selected = excelGrid.SelectedRows.Cast<DataGridViewRow>().Select(r => r.DataBoundItem as ReviewItem).Where(i => i != null).ToList();
            MutateExcel(() => ReviewGroupEditor.RenameGroup(current.Excel, selected, batchGroup.Text));
            BindExcelRows(); RefreshGroupChoices();
            status.Text = "已重命名选中行所属单项工程，含筛选外同组明细；计算表和补差名称已同步。";
        }); }
        void BuildBasisControls(Panel cell, TextBox box) {
            var actions = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 38 };
            actions.Controls.Add(Button("选择模板依据", (s,e) => ChooseBasis(box), 140));
            actions.Controls.Add(Button("恢复模板依据", (s,e) => { string text = Core.Get(current, "模板审核依据"); box.Text = text == "" ? Core.TemplateBasis(Core.SourceTemplate()) : text; }, 140));
            actions.Controls.Add(Button("管理个人依据模板", (s,e) => ManageBasisTemplates(box), 165));
            cell.Height += 40; cell.Controls.Add(actions);
        }
        void ChooseBasis(TextBox box) {
            string basis = Core.Get(current, "模板审核依据"); if (basis == "") basis = Core.TemplateBasis(Core.SourceTemplate());
            basis = BasisChoices(box.Text, basis);
            using (var dialog = new Form { Text = "选择模板审核依据（可继续手动编辑）", Font = Font, Size = new Size(850, 550), StartPosition = FormStartPosition.CenterParent }) {
                var list = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true, HorizontalScrollbar = true };
                var currentLines = box.Text.Replace("\r\n", "\n").Split('\n');
                foreach (string line in basis.Replace("\r\n", "\n").Split('\n').Where(t => !string.IsNullOrWhiteSpace(t))) list.Items.Add(line, currentLines.Contains(line));
                var apply = new Button { Text = "使用勾选条目", Dock = DockStyle.Bottom, Height = 42 };
                apply.Click += (s,e) => { if (list.CheckedItems.Count == 0) { dialog.Text = "请至少勾选一条审核依据"; return; } box.Text = string.Join("\r\n", list.CheckedItems.Cast<string>()); dialog.DialogResult = DialogResult.OK; };
                dialog.Controls.Add(list); dialog.Controls.Add(apply); dialog.ShowDialog(this);
            }
        }
    }
}
