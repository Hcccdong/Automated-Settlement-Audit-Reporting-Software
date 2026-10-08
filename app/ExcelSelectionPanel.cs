using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace SettlementReview {
    public partial class MainForm {
        ReviewItem groupEditItem;
        string groupEditOriginal;
        readonly FlowLayoutPanel excelSelectionBar = new FlowLayoutPanel();
        void BuildExcelSelectionBar() {
            excelSelectionBar.Dock = DockStyle.Top; excelSelectionBar.Height = 30;
            excelSelectionBar.Padding = new Padding(52, 0, 0, 0);
            string[] labels = { "全选", "反选", "全不选" }; int[] modes = { 1, 2, 0 };
            for (int i = 0; i < labels.Length; i++) {
                int mode = modes[i]; var button = new Button { Text = labels[i], Width = i == 2 ? 62 : 48, Height = 25, Font = new Font(Font.FontFamily, 8.5F), Margin = new Padding(0, 1, 5, 1) };
                button.Click += (s,e) => SetVisibleInclude(mode); excelSelectionBar.Controls.Add(button);
            }
            excelSplit.Panel1.Controls.Add(excelSelectionBar); excelGrid.BringToFront();
            excelGrid.Scroll += (s,e) => AlignExcelSelectionBar();
            excelGrid.SizeChanged += (s,e) => AlignExcelSelectionBar();
            excelGrid.CellBeginEdit += (s,e) => { if (e.RowIndex >= 0 && excelGrid.Columns[e.ColumnIndex].DataPropertyName == "Group") { groupEditItem = excelGrid.Rows[e.RowIndex].DataBoundItem as ReviewItem; groupEditOriginal = groupEditItem == null ? null : groupEditItem.Group; } };
            excelGrid.CellEndEdit += (s,e) => {
                if (groupEditItem == null) return;
                var item = groupEditItem; string old = groupEditOriginal, name = item.Group; groupEditItem = null;
                if (name == old) return;
                item.Group = old;
                try { excelLoading = true; ReviewGroupEditor.RenameGroup(current.Excel, new[] { item }, name); }
                catch (Exception ex) { status.Text = ex.Message; }
                finally { excelLoading = false; }
                ExcelChanged(); excelGrid.Refresh(); RefreshGroupChoices();
            };
        }
        void AlignExcelSelectionBar() {
            if (excelGrid.Columns.Count < 2) return;
            int left = excelGrid.GetCellDisplayRectangle(1, -1, true).Left;
            excelSelectionBar.Padding = new Padding(Math.Max(0, left), 0, 0, 0);
        }
    }
}
