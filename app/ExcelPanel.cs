using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SettlementReview {
    public partial class MainForm {
        readonly ReviewGrid excelGrid = new ReviewGrid();
        readonly TextBox excelText = new TextBox(), excelDetails = new TextBox();
        readonly Label excelSummary = new Label(), excelFile = new Label();
        readonly ComboBox excelSheets = new ComboBox(), excelTarget = new ComboBox(), excelStyle = new ComboBox(), excelWording = new ComboBox(), excelDecimals = new ComboBox();
        readonly NumericUpDown excelThreshold = new NumericUpDown();
        readonly CheckBox excelOnlyChanges = new CheckBox();
        readonly CheckBox excelAutoLink = new CheckBox();
        readonly Panel excelBottom = new Panel();
        readonly SplitContainer excelSplit = new SplitContainer();
        Panel excelTop; FlowLayoutPanel excelOptions; TabControl excelPreviews;
        bool excelLoading, excelSplitInitialized;
        void BuildExcelTab() {
            var page = new TabPage("5  Excel对比") { BackColor = Color.White, Padding = new Padding(12) };
            var top = excelTop = new Panel { Dock = DockStyle.Top, Height = 195 };
            var import = Button("导入 Excel 对比表", (s,e) => ImportExcel(), 180); import.Location = new Point(0, 0); top.Controls.Add(import);
            excelConcise.Text = "精简项目特征"; excelConcise.SetBounds(196, 5, 148, 30); excelConcise.CheckedChanged += (s,e) => ExcelChanged(); top.Controls.Add(excelConcise);
            var balances = Button("金额配平", (s,e) => ShowBalanceDialog(), 125); balances.Location = new Point(348, 0); top.Controls.Add(balances);
            excelFile.SetBounds(516, 7, 330, 30); excelFile.AutoEllipsis = true; excelFile.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top; excelFile.Text = "支持 .xlsx / .xlsm，先预览明细，再生成说明。"; top.Controls.Add(excelFile);
            var options = excelOptions = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 153, WrapContents = true };
            options.Controls.Add(new Label { Text = "查看表", AutoSize = true, Margin = new Padding(0, 8, 6, 0) });
            excelSheets.DropDownStyle = ComboBoxStyle.DropDownList; excelSheets.Width = 280; excelSheets.Margin = new Padding(0, 3, 12, 8); excelSheets.SelectedIndexChanged += (s,e) => { if (!excelLoading) BindExcelRows(); }; options.Controls.Add(excelSheets);
            excelOnlyChanges.Text = "仅看有偏差/待核对项"; excelOnlyChanges.Checked = true; excelOnlyChanges.AutoSize = true; excelOnlyChanges.Margin = new Padding(0, 7, 15, 8); excelOnlyChanges.CheckedChanged += (s,e) => BindExcelRows(); options.Controls.Add(excelOnlyChanges);
            options.SetFlowBreak(excelOnlyChanges, true);
            options.Controls.Add(new Label { Text = "列入门槛（元）", AutoSize = true, Margin = new Padding(0, 8, 6, 0) });
            excelThreshold.Width = 105; excelThreshold.Maximum = 1000000000; excelThreshold.DecimalPlaces = 2; excelThreshold.Margin = new Padding(0, 3, 18, 0); excelThreshold.ValueChanged += (s,e) => ExcelChanged(); options.Controls.Add(excelThreshold);
            excelStyle.DropDownStyle = ComboBoxStyle.DropDownList; excelStyle.Items.AddRange(new object[] { "标准说明", "量价明细说明" }); excelStyle.Width = 155; excelStyle.SelectedIndex = 0; excelStyle.SelectedIndexChanged += (s,e) => ExcelChanged(); options.Controls.Add(excelStyle);
            options.Controls.Add(new Label { Text = "金额表述", AutoSize = true, Margin = new Padding(12, 8, 6, 0) });
            excelWording.DropDownStyle = ComboBoxStyle.DropDownList; excelWording.Items.AddRange(new object[] { "正偏差 / 负偏差", "净核减 / 净核增" }); excelWording.Width = 180; excelWording.Margin = new Padding(0, 3, 8, 0); excelWording.SelectedIndexChanged += (s,e) => ExcelChanged(); options.Controls.Add(excelWording);
            options.Controls.Add(new Label { Text = "金额小数位", AutoSize = true, Margin = new Padding(8, 8, 6, 0) });
            excelDecimals.DropDownStyle = ComboBoxStyle.DropDownList; excelDecimals.Items.AddRange(new object[] { "0 位", "1 位", "2 位", "3 位", "4 位" }); excelDecimals.Width = 78; excelDecimals.Margin = new Padding(0, 3, 8, 0); excelDecimals.SelectedIndexChanged += (s,e) => ExcelChanged(); options.Controls.Add(excelDecimals);
            excelAutoLink.Text = "自动联动正文"; excelAutoLink.AutoSize = true; excelAutoLink.Margin = new Padding(8, 6, 0, 0); excelAutoLink.CheckedChanged += (s,e) => { if (!excelLoading && !loading && current.Excel != null) { if (excelAutoLink.Checked) ExcelReview.StartLink(current, current.Excel); else current.Excel.AutoLink = false; ExcelChanged(); } }; options.Controls.Add(excelAutoLink); BuildExcelBatchControls(options); BuildGroupControls(options); top.Controls.Add(options);
            excelGrid.Dock = DockStyle.Fill; excelGrid.AutoGenerateColumns = false; excelGrid.AllowUserToAddRows = false; excelGrid.AllowUserToDeleteRows = false; excelGrid.RowHeadersVisible = false; excelGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect; excelGrid.MultiSelect = true; excelGrid.BackgroundColor = Color.White; excelGrid.BorderStyle = BorderStyle.None;
            excelGrid.EnableHeadersVisualStyles = false; excelGrid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(232, 240, 247); excelGrid.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(31, 56, 78); excelGrid.ColumnHeadersDefaultCellStyle.Font = new Font(Font, FontStyle.Bold); excelGrid.ColumnHeadersHeight = 40; excelGrid.RowTemplate.Height = 36; excelGrid.GridColor = Color.FromArgb(231, 236, 242); excelGrid.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 250, 253); excelGrid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(218, 235, 249); excelGrid.DefaultCellStyle.SelectionForeColor = Color.FromArgb(24, 49, 70);
            excelGrid.Columns.Add(new DataGridViewCheckBoxColumn { DataPropertyName = "Include", HeaderText = "采用", Width = 52, Frozen = true });
            AddExcelColumn("Name", "项目名称", 170, true);
            excelGrid.Columns[1].Frozen = true;
            excelGrid.Columns.Add(new DataGridViewComboBoxColumn { DataPropertyName = "Reason", HeaderText = "调整原因 / 文字模板", Width = 205, DataSource = ExcelReview.Reasons });
            AddExcelColumn("SentAmount", "送审（元）", 110, true); AddExcelColumn("AuditAmount", "审查（元）", 110, true);
            AddExcelColumn("IncreaseAmount", "核增（元）", 110, true); AddExcelColumn("DecreaseAmount", "核减（元）", 110, true);
            AddExcelColumn("Direction", "偏差方向", 155, true); AddExcelColumn("Delta", "偏差金额（元）", 145, true);
            AddExcelColumn("Feature", "项目特征（完整内容见下方）", 230, true);
            AddExcelColumn("Group", "单项工程（可批量覆盖）", 190, false);
            excelGrid.Columns.Add(new DataGridViewComboBoxColumn { DataPropertyName = "Target", HeaderText = "报告归属", Width = 205, DataSource = ExcelReview.Targets });
            AddExcelColumn("Note", "补充原因（可填写）", 200, false);
            excelGrid.Columns.Add(new DataGridViewCheckBoxColumn { DataPropertyName = "Confirmed", HeaderText = "提示已核实", Width = 100 });
            AddExcelColumn("Warning", "数据提示", 320, true); AddExcelColumn("Location", "原表位置", 260, true);
            excelGrid.CurrentCellDirtyStateChanged += (s,e) => { if (excelGrid.IsCurrentCellDirty) excelGrid.CommitEdit(DataGridViewDataErrorContexts.Commit); };
            excelGrid.CellValueChanged += (s,e) => ExcelChanged(); excelGrid.SelectionChanged += (s,e) => ShowExcelDetails();
            excelGrid.CellFormatting += (s,e) => { if (e.RowIndex < 0) return; var item = excelGrid.Rows[e.RowIndex].DataBoundItem as ReviewItem; if (item != null && excelGrid.Columns[e.ColumnIndex].DataPropertyName == "Direction") { e.Value = ExcelReview.DirectionText(current.Excel, item); e.FormattingApplied = true; } if (item != null && item.Warning != "" && !item.Confirmed) e.CellStyle.ForeColor = Color.FromArgb(157, 80, 18); };
            SetupExcelBatchEvents();
            excelGrid.DataError += (s,e) => { e.ThrowException = false; status.Text = "此单元格内容无效，请重新选择或填写。"; };
            var bottom = excelBottom; bottom.Dock = DockStyle.Fill; bottom.BackColor = Color.White;
            excelSummary.Dock = DockStyle.Top; excelSummary.Height = 78; excelSummary.Padding = new Padding(3, 6, 0, 0); excelSummary.ForeColor = Color.FromArgb(31, 85, 119); bottom.Controls.Add(excelSummary);
            var previews = excelPreviews = new TabControl { Dock = DockStyle.Fill }; var textTab = new TabPage("将写入报告的文字（自动更新）"); var detailTab = new TabPage("所选项目及导入提示");
            foreach (var box in new[] { excelText, excelDetails }) { box.Multiline = true; box.ReadOnly = true; box.ScrollBars = ScrollBars.Both; box.Dock = DockStyle.Fill; box.BackColor = Color.FromArgb(247, 250, 253); box.BorderStyle = BorderStyle.FixedSingle; }
            textTab.Controls.Add(excelText); detailTab.Controls.Add(excelDetails); previews.TabPages.Add(textTab); previews.TabPages.Add(detailTab); bottom.Controls.Add(previews); previews.BringToFront();
            var apply = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 48, Padding = new Padding(0, 7, 0, 0) };
            excelTarget.DropDownStyle = ComboBoxStyle.DropDownList; excelTarget.Items.AddRange(ExcelReview.Targets); excelTarget.SelectedIndex = 0; excelTarget.Width = 190; apply.Controls.Add(excelTarget);
            apply.Controls.Add(Button("所查看表归入此处", (s,e) => AssignExcelTarget(), 175));
            apply.Controls.Add(Button("查看联动正文", (s,e) => ShowLinkedBody(), 140)); bottom.Controls.Add(apply);
            excelSplit.Dock = DockStyle.Fill; excelSplit.Orientation = Orientation.Horizontal; excelSplit.SplitterWidth = 7; excelSplit.FixedPanel = FixedPanel.Panel2; excelSplit.Panel1MinSize = 190; excelSplit.Panel2MinSize = 118; excelSplit.BackColor = Color.FromArgb(205, 216, 225);
            excelSplit.Panel1.Controls.Add(excelGrid); excelSplit.Panel2.Controls.Add(bottom); BuildExcelSelectionBar();
            page.Controls.Add(excelSplit); page.Controls.Add(top); tabs.TabPages.Add(page);
            Resize += (s,e) => AdjustExcelLayout();
        }
        void AdjustExcelLayout() {
            float scale = DeviceDpi / 96F; bool compact = ClientSize.Height / scale < 760;
            excelSummary.Height = (int)((compact ? 56 : 78) * scale); if (excelPreviews != null) excelPreviews.Visible = true;
            if (excelTop != null) { excelTop.Height = (int)(195 * scale); excelOptions.Height = (int)(153 * scale); excelFile.Width = Math.Max(0, excelTop.ClientSize.Width - excelFile.Left - (int)(8 * scale)); }
            if (!excelSplitInitialized && excelSplit.Height > excelSplit.Panel1MinSize + excelSplit.Panel2MinSize + excelSplit.SplitterWidth) {
                int desiredBottom = (int)((compact ? 145 : 180) * scale);
                int distance = Math.Max(excelSplit.Panel1MinSize, Math.Min(excelSplit.Height - excelSplit.Panel2MinSize - excelSplit.SplitterWidth, excelSplit.Height - desiredBottom - excelSplit.SplitterWidth));
                excelSplit.SplitterDistance = distance; excelSplitInitialized = true;
            }
            UpdateExcelSummary();
        }
        void AddExcelColumn(string field, string title, int width, bool readOnly) {
            var col = new DataGridViewTextBoxColumn { DataPropertyName = field, HeaderText = title, Width = width, ReadOnly = readOnly, SortMode = DataGridViewColumnSortMode.NotSortable };
            if (field.Contains("Amount") || field == "Delta") { col.DefaultCellStyle.Format = "0.00"; col.DefaultCellStyle.NullValue = "未提供"; col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; }
            excelGrid.Columns.Add(col);
        }
        void ImportExcel() {
            using (var dialog = new OpenFileDialog { Filter = "Excel 对比表 (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|旧版 Excel (*.xls)|*.xls", Title = "导入审核对比表（金额单位：元）" }) if (dialog.ShowDialog(this) == DialogResult.OK) Run(() => {
                UseWaitCursor = true;
                try { var parsed = ExcelReview.Parse(ExcelReview.Read(dialog.FileName), Path.GetFileName(dialog.FileName)); if (parsed.Items.Count == 0) throw new InvalidDataException(string.Join("\n", parsed.Messages)); parsed.WordingMode = ExcelReview.DetectWordingMode(CurrentWordingTemplate()); Gather(); BackupBeforeTemplate(); ExcelReview.StartLink(current, parsed); LoadExcelState(); Changed(); status.Text = "已导入并联动正文；金额表述已按模板预选，可手动切换。"; }
                finally { UseWaitCursor = false; }
            });
        }
        void LoadExcelState() {
            excelLoading = true; excelProblemsOnly.Checked = false;
            if (current.Excel == null) current.Excel = new ExcelReviewState();
            if (current.Excel.Items.Count > 0 && current.Excel.ModelVersion < 6) ExcelReview.StartLink(current, current.Excel);
            ExcelReview.Normalize(current.Excel); excelConcise.Checked = current.Excel.CompactFeatures; excelAutoLink.Checked = current.Excel.AutoLink; excelWording.SelectedIndex = current.Excel.WordingMode == ExcelReview.NetWording ? 1 : 0; excelDecimals.SelectedIndex = current.Excel.AmountDecimals;
            excelSheets.Items.Clear(); excelSheets.Items.Add("全部工作表"); foreach (string sheet in current.Excel.Items.Select(i => i.Sheet).Distinct()) excelSheets.Items.Add(sheet); excelSheets.SelectedIndex = 0;
            excelThreshold.Value = Math.Min(excelThreshold.Maximum, Math.Max(0, current.Excel.Threshold)); excelStyle.SelectedIndex = current.Excel.Detailed ? 1 : 0;
            excelFile.Text = current.Excel.FileName == "" ? "支持 .xlsx / .xlsm。旧式 .xls 请先另存为 .xlsx。" : current.Excel.FileName + "  ·  " + current.Excel.Items.Count + " 条明细";
            excelLoading = false; ApplyExcelDecimalFormat(); BindExcelRows(); UpdateExcelPreview();
        }
        void UpdateLinkedBody() {
            if (current.Excel == null) return;
            ExcelReview.Synchronize(current, false); bool wasLoading = loading; loading = true;
            try { foreach (string field in ExcelReview.Targets) if (inputs.ContainsKey(field)) { inputs[field].ReadOnly = current.Excel.AutoLink; inputs[field].BackColor = current.Excel.AutoLink ? Color.FromArgb(244, 248, 252) : Color.White; if (current.Excel.AutoLink && inputs[field].Text != Core.Get(current, field)) inputs[field].Text = Core.Get(current, field); } }
            finally { loading = wasLoading; }
        }
        void ShowLinkedBody() {
            Changed(); tabs.SelectedIndex = 3;
            string field = excelTarget.SelectedItem == null ? ExcelReview.Targets[0] : excelTarget.SelectedItem.ToString();
            if (!inputs.ContainsKey(field)) return;
            Control box = inputs[field], ancestor = box.Parent;
            while (ancestor != null) { var scroll = ancestor as ScrollableControl; if (scroll != null && scroll.AutoScroll) { var location = scroll.PointToClient(box.PointToScreen(Point.Empty)); scroll.AutoScrollPosition = new Point(0, Math.Max(0, location.Y - scroll.AutoScrollPosition.Y - 16)); break; } ancestor = ancestor.Parent; }
            inputs[field].SelectionLength = 0;
        }
        void AssignExcelTarget() { Run(() => {
            if (current.Excel == null || current.Excel.Items.Count == 0) throw new InvalidDataException("请先导入 Excel 对比表。");
            excelGrid.EndEdit(); string sheet = excelSheets.SelectedIndex > 0 ? excelSheets.SelectedItem.ToString() : null;
            foreach (var item in current.Excel.Items.Where(i => sheet == null || i.Sheet == sheet)) item.Target = excelTarget.SelectedItem.ToString();
            foreach (var balance in current.Excel.Balances) {
                var targets = current.Excel.Items.Where(i => ExcelReview.Norm(i.Group) == ExcelReview.Norm(balance.Group)).Select(i => i.Target).Distinct().ToList();
                if (targets.Count == 1) balance.Target = targets[0];
            }
            BindExcelRows(); ExcelChanged(); status.Text = "已更新所查看工作表的报告归属，正文同步更新。";
        }); }
        void BindExcelRows() {
            if (excelLoading || current.Excel == null) return; CancelReasonFill(); includeAnchor = -1; excelLoading = true;
            string sheet = excelSheets.SelectedIndex > 0 ? excelSheets.SelectedItem.ToString() : null;
            var problems = new HashSet<ReviewItem>(excelProblemsOnly.Checked ? ExcelReview.ProblemRows(current.Excel) : new List<ReviewItem>());
            excelGrid.DataSource = new BindingList<ReviewItem>(current.Excel.Items.Where(i => (sheet == null || i.Sheet == sheet) && (!excelOnlyChanges.Checked || i.Delta != 0 || i.Warning != "") && (!excelProblemsOnly.Checked || problems.Contains(i))).ToList());
            foreach (DataGridViewRow row in excelGrid.Rows) row.Height = (int)(36 * DeviceDpi / 96F);
            excelLoading = false; RefreshExcelIssues(); ShowExcelDetails();
        }
        void ExcelChanged() {
            if (loading || excelLoading || current.Excel == null) return;
            current.Excel.Threshold = excelThreshold.Value; current.Excel.Detailed = excelStyle.SelectedIndex == 1; current.Excel.CompactFeatures = excelConcise.Checked; current.Excel.WordingMode = excelWording.SelectedIndex == 1 ? ExcelReview.NetWording : ExcelReview.DirectionWording; current.Excel.AmountDecimals = excelDecimals.SelectedIndex < 0 ? 2 : excelDecimals.SelectedIndex;
            ApplyExcelDecimalFormat(); Changed(); UpdateExcelPreview(); RefreshExcelIssues(); excelGrid.Invalidate();
        }
        void ApplyExcelDecimalFormat() {
            if (current.Excel == null) return; string format = "F" + current.Excel.AmountDecimals;
            foreach (DataGridViewColumn col in excelGrid.Columns) if (new[] { "SentAmount", "AuditAmount", "IncreaseAmount", "DecreaseAmount", "Delta" }.Contains(col.DataPropertyName)) col.DefaultCellStyle.Format = format;
        }
        void UpdateExcelPreview() {
            if (current.Excel == null) return;
            UpdateExcelSummary(); excelText.Text = string.Join("\r\n\r\n", ExcelReview.Targets.Select(t => new { Target = t, Text = ExcelReview.Compose(current.Excel, false, t) }).Where(t => t.Text != "无").Select(t => t.Target + "\r\n" + t.Text));
            if (excelText.Text == "") excelText.Text = current.Excel.Items.Count == 0 ? "先导入 Excel，再选择采用项和调整原因。" : "尚未选择可组稿的偏差项。可取消“仅看有偏差/待核对项”查看所有明细。";
            ShowExcelDetails();
        }
        void UpdateExcelSummary() {
            if (current.Excel == null) return; bool compact = ClientSize.Height / (DeviceDpi / 96F) < 760;
            string summary = ExcelReview.Summary(current.Excel); int balanced = current.Excel.Balances == null ? 0 : current.Excel.Balances.Count(b => b.Enabled); if (balanced > 0) summary = summary.Split('\n')[0].TrimEnd() + "\n" + balanced + " 个单项工程启用末条补差，点击“金额配平”核对总额。"; if (compact) summary = summary.Split('\n')[0].TrimEnd();
            excelSummary.Text = summary + "\n待处理 " + ExcelReview.ProblemRows(current.Excel).Count + " 项（红色） · " + (compact ? "拖动下方分隔条查看自动组稿文字" : "Shift 连选行；原因格右下角蓝点可拖拽填充");
        }
        void ShowExcelDetails() {
            if (current.Excel == null) return;
            var item = excelGrid.CurrentRow == null ? null : excelGrid.CurrentRow.DataBoundItem as ReviewItem;
            excelDetails.Text = (item == null ? "" : item.Location + "\r\n单项工程：" + item.Group + "  /  " + item.Target + "\r\n送审：" + item.Code + "  " + item.Name + "\r\n项目特征：" + item.Feature + "\r\n工程量：" + ExcelReview.Number(item.SentQuantity) + " " + item.Unit + "  单价：" + ExcelReview.Number(item.SentPrice) + " 元\r\n审查：" + item.AuditCode + "  " + item.AuditName + "\r\n审查特征：" + item.AuditFeature + "\r\n工程量：" + ExcelReview.Number(item.AuditQuantity) + " " + item.AuditUnit + "  单价：" + ExcelReview.Number(item.AuditPrice) + " 元\r\n原表核增：" + ExcelReview.Money(current.Excel, item.SourceIncrease) + "  原表核减：" + ExcelReview.Money(current.Excel, item.SourceDecrease) + "\r\n" + item.Warning + "\r\n待处理：" + ActiveRowIssue(item) + "\r\n\r\n") + string.Join("\r\n", current.Excel.Messages);
        }
        void ApplyExcelText() { Run(() => {
            excelGrid.EndEdit(); Gather(); string target = excelTarget.SelectedItem.ToString();
            var keys = Core.ValidateTemplate(CurrentTemplate());
            if (!keys.Contains(target) && !keys.Contains("修正造价情况说明")) throw new InvalidDataException("当前 Word 模板尚未设置“" + target + "”填写位置。请到第一步关联后再写入。");
            string text = ExcelReview.Compose(current.Excel, true); ExcelReview.Apply(current, target, text); inputs[target].Text = current.Values[target]; Changed(); status.Text = "已写入“" + target + "”，可在审核正文与报告导出页查看。";
        }); }
    }
}
