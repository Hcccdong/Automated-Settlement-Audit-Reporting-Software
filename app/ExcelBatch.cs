using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace SettlementReview {
    public class ReviewGrid : DataGridView {
        public Func<MouseEventArgs, bool> FillMouseDown;
        public Action<MouseEventArgs> FillMouseUp;
        bool customPress;
        protected override void OnMouseDown(MouseEventArgs e) { customPress = FillMouseDown != null && FillMouseDown(e); if (customPress) return; base.OnMouseDown(e); }
        protected override void OnMouseUp(MouseEventArgs e) { if (FillMouseUp != null) FillMouseUp(e); if (customPress) { customPress = false; Capture = false; return; } base.OnMouseUp(e); }
    }
    public partial class MainForm {
        readonly CheckBox excelProblemsOnly = new CheckBox();
        readonly ComboBox batchReason = new ComboBox();
        readonly Timer fillScroll = new Timer();
        int includeAnchor = -1, fillStart = -1, fillEnd = -1;
        string fillReason = "";
        bool reasonDragging;
        int ExcelColumn(string property) { return excelGrid.Columns.Cast<DataGridViewColumn>().First(c => c.DataPropertyName == property).Index; }
        List<ReviewItem> VisibleExcelItems() { return excelGrid.Rows.Cast<DataGridViewRow>().Select(r => r.DataBoundItem as ReviewItem).Where(i => i != null).ToList(); }
        void MutateExcel(Action change) {
            excelGrid.EndEdit(); excelLoading = true;
            try { change(); } finally { excelLoading = false; }
            excelGrid.Refresh(); ExcelChanged();
        }
        void SetVisibleInclude(int mode) {
            MutateExcel(() => { foreach (var item in VisibleExcelItems()) item.Include = mode == 0 ? false : mode == 1 ? true : !item.Include; });
            status.Text = "已更新当前显示的 " + excelGrid.Rows.Count + " 行采用状态；其他筛选外明细保留。";
        }
        void IncludeRange(int first, int last, bool value) {
            var rows = VisibleExcelItems();
            MutateExcel(() => { for (int i = Math.Max(0, Math.Min(first, last)); i <= Math.Min(rows.Count - 1, Math.Max(first, last)); i++) rows[i].Include = value; });
        }
        bool HandleIncludeClick(Point position, bool shift) {
            var hit = excelGrid.HitTest(position.X, position.Y);
            if (hit.RowIndex < 0 || hit.ColumnIndex < 0 || excelGrid.Columns[hit.ColumnIndex].DataPropertyName != "Include") return false;
            excelGrid.EndEdit(); var clicked = (ReviewItem)excelGrid.Rows[hit.RowIndex].DataBoundItem;
            if (shift && includeAnchor >= 0 && includeAnchor < excelGrid.Rows.Count) {
                IncludeRange(includeAnchor, hit.RowIndex, ((ReviewItem)excelGrid.Rows[includeAnchor].DataBoundItem).Include);
                return true;
            }
            var selected = excelGrid.SelectedRows.Cast<DataGridViewRow>().ToList();
            if (!excelGrid.Rows[hit.RowIndex].Selected || selected.Count == 0) {
                excelGrid.CurrentCell = excelGrid.Rows[hit.RowIndex].Cells[hit.ColumnIndex]; excelGrid.ClearSelection(); excelGrid.Rows[hit.RowIndex].Selected = true;
                selected = new List<DataGridViewRow> { excelGrid.Rows[hit.RowIndex] };
            }
            bool value = !clicked.Include;
            MutateExcel(() => { foreach (var row in selected) ((ReviewItem)row.DataBoundItem).Include = value; });
            includeAnchor = hit.RowIndex;
            status.Text = "已" + (value ? "勾选" : "取消勾选") + "选中的 " + selected.Count + " 项。";
            return true;
        }
        void FillReasons(int first, int last, string reason) {
            if (!ExcelReview.Reasons.Contains(reason)) return;
            var rows = VisibleExcelItems();
            MutateExcel(() => { for (int i = Math.Max(0, Math.Min(first, last)); i <= Math.Min(rows.Count - 1, Math.Max(first, last)); i++) rows[i].Reason = reason; });
            status.Text = "已填充调整原因。补充说明和采用状态保留，请逐项核对适用性。";
        }
        void ApplySelectedReason() {
            if (batchReason.SelectedItem == null) return;
            var selected = excelGrid.SelectedRows.Cast<DataGridViewRow>().Select(r => r.DataBoundItem as ReviewItem).Where(i => i != null).ToList();
            MutateExcel(() => { foreach (var item in selected) item.Reason = batchReason.SelectedItem.ToString(); });
            status.Text = "已将调整原因应用到选中的 " + selected.Count + " 行。";
        }
        void BuildExcelBatchControls(FlowLayoutPanel options) {
            options.SetFlowBreak(excelAutoLink, true);
            batchReason.DropDownStyle = ComboBoxStyle.DropDownList; batchReason.Width = 145; batchReason.Items.AddRange(ExcelReview.Reasons); batchReason.SelectedIndex = 0; options.Controls.Add(batchReason);
            options.Controls.Add(Button("应用到选中行", (s,e) => ApplySelectedReason(), 125));
            excelProblemsOnly.Text = "仅看待处理"; excelProblemsOnly.AutoSize = true; excelProblemsOnly.Margin = new Padding(4, 7, 8, 0); excelProblemsOnly.CheckedChanged += (s,e) => BindExcelRows(); options.Controls.Add(excelProblemsOnly);
            options.Controls.Add(Button("定位待处理", (s,e) => LocateExcelProblems(), 104));
        }
        bool LocateExcelProblems() {
            if (current.Excel == null) return false;
            excelGrid.EndEdit(); var problems = ExcelReview.ProblemRows(current.Excel);
            if (problems.Count == 0) { status.Text = "Excel 已采用明细没有待处理项。"; return false; }
            var previous = excelGrid.CurrentRow == null ? null : excelGrid.CurrentRow.DataBoundItem as ReviewItem;
            var target = problems[tabs.SelectedIndex == 4 ? (Math.Max(-1, problems.IndexOf(previous)) + 1) % problems.Count : 0];
            excelLoading = true; excelSheets.SelectedIndex = 0; excelOnlyChanges.Checked = false; excelProblemsOnly.Checked = true; excelLoading = false;
            BindExcelRows(); tabs.SelectedIndex = 4;
            var row = excelGrid.Rows.Cast<DataGridViewRow>().First(r => ReferenceEquals(r.DataBoundItem, target));
            string key = !ExcelReview.Reasons.Contains(target.Reason) || target.Reason == "待确认" ? "Reason" : target.Reason == ExcelReview.Reasons[3] && string.IsNullOrWhiteSpace(target.Note) ? "Note" : string.IsNullOrWhiteSpace(target.Group) || target.Group == "未填写工程名称" ? "Group" : "Confirmed";
            excelGrid.ClearSelection(); excelGrid.CurrentCell = row.Cells[ExcelColumn(key)]; row.Selected = true; excelGrid.FirstDisplayedScrollingRowIndex = row.Index; excelGrid.Focus();
            status.Text = "有 " + problems.Count + " 项待处理，已标红并定位：" + target.Location + " · " + ExcelReview.RowIssue(target);
            return true;
        }
        string ActiveRowIssue(ReviewItem item) {
            if (item == null || !item.Include || item.Delta.HasValue && (item.Delta == 0 || Math.Abs(item.Delta.Value) < current.Excel.Threshold)) return "";
            return ExcelReview.RowIssue(item);
        }
        void RefreshExcelIssues() {
            foreach (DataGridViewRow row in excelGrid.Rows) {
                var item = row.DataBoundItem as ReviewItem; string issue = ActiveRowIssue(item); row.ErrorText = issue;
                foreach (DataGridViewCell cell in row.Cells) cell.ToolTipText = issue;
            }
            excelGrid.Invalidate();
            if (status.Text.StartsWith("有 ") && status.Text.Contains("项待处理")) status.Text = "Excel 待处理 " + ExcelReview.ProblemRows(current.Excel).Count + " 项；已修正行的红色提示已清除。";
        }
        void ShowExcelTextPreview() {
            UpdateExcelPreview(); if (excelPreviews.Visible) { excelPreviews.SelectedIndex = 0; return; }
            using (var dialog = new Form { Text = "Excel 生成文字预览", Font = Font, StartPosition = FormStartPosition.CenterParent, Size = new Size(850, 600), MinimizeBox = false }) {
                dialog.Controls.Add(new TextBox { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, Text = excelText.Text, BackColor = Color.White }); dialog.ShowDialog(this);
            }
        }
        Rectangle ReasonFillHandle() {
            var cell = excelGrid.CurrentCell;
            if (cell == null || cell.RowIndex < 0 || excelGrid.Columns[cell.ColumnIndex].DataPropertyName != "Reason") return Rectangle.Empty;
            Rectangle rect = excelGrid.GetCellDisplayRectangle(cell.ColumnIndex, cell.RowIndex, true);
            int size = Math.Max(8, (int)(8 * DeviceDpi / 96F));
            return rect.Width < size || rect.Height < size ? Rectangle.Empty : new Rectangle(rect.Right - size - 2, rect.Bottom - size - 2, size, size);
        }
        void TrackReasonFill(Point mouse) {
            if (!reasonDragging) return;
            var hit = excelGrid.HitTest(mouse.X, mouse.Y);
            if (hit.RowIndex >= 0) fillEnd = hit.RowIndex;
            excelGrid.Invalidate();
        }
        void CancelReasonFill() { reasonDragging = false; fillScroll.Stop(); excelGrid.Capture = false; excelGrid.Invalidate(); }
        void SetupExcelBatchEvents() {
            excelGrid.CellFormatting += (s,e) => {
                if (e.RowIndex < 0) return;
                var item = excelGrid.Rows[e.RowIndex].DataBoundItem as ReviewItem; string issue = ActiveRowIssue(item);
                if (issue != "") { e.CellStyle.BackColor = Color.FromArgb(255, 232, 229); e.CellStyle.ForeColor = Color.FromArgb(155, 36, 32); e.CellStyle.SelectionBackColor = Color.FromArgb(250, 202, 195); e.CellStyle.SelectionForeColor = Color.FromArgb(112, 24, 20); }
                if (reasonDragging && e.RowIndex >= Math.Min(fillStart, fillEnd) && e.RowIndex <= Math.Max(fillStart, fillEnd) && excelGrid.Columns[e.ColumnIndex].DataPropertyName == "Reason") e.CellStyle.BackColor = Color.FromArgb(191, 224, 243);
            };
            excelGrid.CellPainting += (s,e) => {
                if (e.RowIndex < 0 || excelGrid.CurrentCell == null || e.RowIndex != excelGrid.CurrentCell.RowIndex || e.ColumnIndex != excelGrid.CurrentCell.ColumnIndex || excelGrid.Columns[e.ColumnIndex].DataPropertyName != "Reason") return;
                e.Paint(e.ClipBounds, DataGridViewPaintParts.All); using (var brush = new SolidBrush(Color.FromArgb(22, 102, 163))) e.Graphics.FillRectangle(brush, ReasonFillHandle()); e.Handled = true;
            };
            excelGrid.FillMouseDown = e => {
                if (e.Button == MouseButtons.Left && HandleIncludeClick(e.Location, (ModifierKeys & Keys.Shift) != 0)) return true;
                if (e.Button != MouseButtons.Left || !ReasonFillHandle().Contains(e.Location)) return false;
                excelGrid.EndEdit(); fillStart = excelGrid.CurrentCell.RowIndex; fillEnd = fillStart;
                fillReason = ((ReviewItem)excelGrid.Rows[fillStart].DataBoundItem).Reason;
                reasonDragging = true; excelGrid.Capture = true; fillScroll.Start(); return true;
            };
            excelGrid.MouseMove += (s,e) => { if (reasonDragging) TrackReasonFill(e.Location); excelGrid.Cursor = reasonDragging || ReasonFillHandle().Contains(e.Location) ? Cursors.Cross : Cursors.Default; };
            excelGrid.FillMouseUp = e => { if (!reasonDragging) return; TrackReasonFill(e.Location); int first = fillStart, last = fillEnd; string reason = fillReason; CancelReasonFill(); FillReasons(first, last, reason); };
            excelGrid.KeyDown += (s,e) => { if (e.KeyCode == Keys.Escape && reasonDragging) { CancelReasonFill(); e.Handled = true; } };
            excelGrid.MouseCaptureChanged += (s,e) => { if (!excelGrid.Capture && reasonDragging) CancelReasonFill(); };
            fillScroll.Interval = 100; fillScroll.Tick += (s,e) => {
                if (!reasonDragging || excelGrid.Rows.Count == 0) return;
                Point pos = excelGrid.PointToClient(Cursor.Position); int first = excelGrid.FirstDisplayedScrollingRowIndex;
                if (first < 0) return;
                if (pos.Y > excelGrid.ClientSize.Height - 30 && first < excelGrid.Rows.Count - 1) excelGrid.FirstDisplayedScrollingRowIndex = first + 1;
                else if (pos.Y < excelGrid.ColumnHeadersHeight + 20 && first > 0) excelGrid.FirstDisplayedScrollingRowIndex = first - 1;
                TrackReasonFill(new Point(Math.Max(60, pos.X), Math.Max(excelGrid.ColumnHeadersHeight + 2, Math.Min(excelGrid.ClientSize.Height - 24, pos.Y))));
            };
            FormClosed += (s,e) => { fillScroll.Dispose(); };
        }
    }
}
