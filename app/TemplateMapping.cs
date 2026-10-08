using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;

namespace SettlementReview {
    public class TemplateParagraph {
        public string Part, Text; public int Index;
        public override string ToString() {
            string place = Part.Contains("header") ? "页眉" : Part.Contains("footer") ? "页脚" : Part == "word/document.xml" ? "正文" : "其他";
            return place + " " + (Index + 1) + "  " + (Text.Length > 70 ? Text.Substring(0, 70) + "…" : Text);
        }
    }
    public static partial class Core {
        public static List<TemplateParagraph> TemplateParagraphs(byte[] bytes) {
            var result = new List<TemplateParagraph>();
            using (var ms = new MemoryStream(bytes)) using (var z = new ZipArchive(ms, ZipArchiveMode.Read)) {
                foreach (var e in z.Entries.Where(e => IsTextPart(e.FullName))) {
                    int i = 0; foreach (var p in Xml(e).Descendants(W + "p")) {
                        string text = RawParagraphText(p);
                        if (!string.IsNullOrWhiteSpace(text)) result.Add(new TemplateParagraph { Part = e.FullName, Index = i, Text = text }); i++;
                    }
                }
            } return result.OrderBy(x => x.Part == "word/document.xml" ? 0 : x.Part.Contains("header") ? 1 : x.Part.Contains("footer") ? 2 : 3).ThenBy(x => x.Part).ThenBy(x => x.Index).ToList();
        }
        public static bool HasTemplateFields(byte[] bytes) { return TemplateParagraphs(bytes).Any(p => Regex.IsMatch(p.Text, @"\{\{[^{}]+\}\}")); }
        public static byte[] MapSelection(byte[] bytes, TemplateParagraph selected, int start, int length, string field) {
            if (!Allowed.Contains(field)) throw new InvalidDataException("请选择填写项。");
            using (var input = new MemoryStream(bytes)) using (var z = new ZipArchive(input, ZipArchiveMode.Read)) using (var output = new MemoryStream()) {
                using (var target = new ZipArchive(output, ZipArchiveMode.Create, true)) {
                    foreach (var e in z.Entries) using (var dest = target.CreateEntry(e.FullName).Open()) {
                        if (e.FullName == selected.Part) {
                            var xml = Xml(e); var p = xml.Descendants(W + "p").ElementAt(selected.Index);
                            string text = RawParagraphText(p);
                            if (text != selected.Text) throw new InvalidDataException("该段文字已更新，请重新选择。");
                            foreach (Match m in Regex.Matches(text, @"\{\{[^{}]+\}\}"))
                                if (start < m.Index + m.Length && start + length > m.Index && (start > m.Index || start + length < m.Index + m.Length)) throw new InvalidDataException("如需修改已关联项，请选中完整的字段标记。");
                            ReplaceSpan(p, start, length, "{{" + field + "}}");
                            using (var writer = new StreamWriter(dest, new UTF8Encoding(false))) xml.Save(writer, SaveOptions.DisableFormatting);
                        } else using (var source = e.Open()) source.CopyTo(dest);
                    }
                } return output.ToArray();
            }
        }
        public static IEnumerable<string> RequiredForTemplate(IEnumerable<string> keys) {
            var result = new HashSet<string>(keys);
            if (result.Overlaps(new[] { "审核结论", "金额汇总说明", "净核减金额", "净核减率" })) result.UnionWith(new[] { "送审金额", "审定金额", "核减合计", "审增合计" });
            if (result.Overlaps(new[] { "审核结论", "金额汇总说明" })) result.Add("净核减金额");
            if (result.Contains("审核金额")) result.Add("审定金额");
            if (result.Contains("核减金额")) result.Add("核减合计");
            if (result.Contains("核增金额")) result.Add("审增合计");
            if (result.Contains("有关事项说明")) result.UnionWith(MatterFields);
            if (result.Contains("签发日期")) result.Add("报告日期");
            if (result.Contains("修正造价情况说明")) result.UnionWith(new[] { "合同标内调整说明", "合同标外调整说明", "联系单调整说明", "其他调整说明" });
            if (result.Contains("抄送单位")) result.UnionWith(new[] { "建设单位", "施工单位" });
            if (result.Contains("工期情况")) result.UnionWith(Core.DateFields.Concat(new[] { "合同工期日历天数" }));
            if (result.Contains("实际工期日历天数")) result.UnionWith(new[] { "实际开工", "实际竣工" });
            if (result.Contains("计划工期日历天数")) result.UnionWith(new[] { "计划开工", "计划竣工" });
            if (result.Contains("工期延期天数")) result.UnionWith(new[] { "实际开工", "实际竣工", "合同工期日历天数" });
            return result;
        }
        public static string ConclusionPreview(Project p, byte[] template) {
            var values = Values(p, true);
            var ps = TemplateParagraphs(PrepareTemplate(template)).Where(x => x.Part == "word/document.xml").ToList();
            int heading = ps.FindIndex(x => Regex.Replace(x.Text, @"^[一二三四五六七八九十\d\s、.．（）()]+", "").Trim() == "审核结论");
            string text = heading >= 0 && heading + 1 < ps.Count ? ps[heading + 1].Text : ps.Where(x => x.Text.Contains("{{审核结论}}")).Select(x => x.Text).FirstOrDefault();
            return text == null ? "当前模板未单独设置审核结论段，请在模板预览中查看。" : Regex.Replace(text, @"\{\{([^{}]+)\}\}", m => values.ContainsKey(m.Groups[1].Value) ? values[m.Groups[1].Value] : m.Value);
        }
    }
    public class TemplateMappingForm : Form {
        public byte[] Result { get; private set; }
        readonly ListBox paragraphs = new ListBox(); readonly TextBox text = new TextBox(); readonly ComboBox fields = new ComboBox(); readonly Label status = new Label();
        readonly Stack<byte[]> undo = new Stack<byte[]>();
        public TemplateMappingForm(byte[] source) {
            Result = (byte[])source.Clone(); Text = "设置模板填写位置"; Size = new Size(1040, 760); MinimumSize = new Size(900, 640); Font = new Font("Microsoft YaHei UI", 10.5F); StartPosition = FormStartPosition.CenterParent; AutoScaleMode = AutoScaleMode.None;
            var help = new Label { Dock = DockStyle.Top, Height = 72, Padding = new Padding(14), Text = "选择一个段落，在下方选中需要替换的文字，再关联右侧填写项。\n未关联的文字将按模板原样保留。这里仅显示文字；下划线、横线和版式仍保存在原文档中。" };
            var split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal, SplitterDistance = 265 };
            paragraphs.Dock = DockStyle.Fill; paragraphs.HorizontalScrollbar = true; split.Panel1.Controls.Add(paragraphs);
            text.Dock = DockStyle.Fill; text.Multiline = true; text.ReadOnly = true; text.BackColor = Color.White; text.ScrollBars = ScrollBars.Both; text.HideSelection = false; text.WordWrap = true; split.Panel2.Controls.Add(text);
            var bar = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 58, Padding = new Padding(12, 8, 8, 8) };
            fields.DropDownStyle = ComboBoxStyle.DropDownList; fields.Width = 190; fields.Items.AddRange(Core.Allowed.OrderBy(s => s).Cast<object>().ToArray()); fields.SelectedItem = "项目名称"; bar.Controls.Add(fields);
            AddButton(bar, "关联选中文字", () => {
                try {
                    var p = paragraphs.SelectedItem as TemplateParagraph; if (p == null || text.SelectionLength == 0) throw new InvalidDataException("请先在文字框中选中需要替换的文字。");
                    byte[] next = Core.MapSelection(Result, p, text.SelectionStart, text.SelectionLength, fields.Text);
                    undo.Push(Result); Result = next; RefreshList(p.Part, p.Index); status.Text = "已关联为“" + fields.Text + "”；其余文字和格式保留。";
                } catch (Exception ex) { status.Text = ex.Message; }
            });
            AddButton(bar, "撤销", () => { if (undo.Count > 0) { Result = undo.Pop(); RefreshList(null, 0); } });
            AddButton(bar, "保存并使用", () => { try { Core.ValidateTemplate(Result); DialogResult = DialogResult.OK; Close(); } catch (Exception ex) { status.Text = ex.Message; } });
            AddButton(bar, "取消", () => { DialogResult = DialogResult.Cancel; Close(); });
            status.Dock = DockStyle.Bottom; status.Height = 40; status.Padding = new Padding(12, 6, 8, 4); status.ForeColor = Color.FromArgb(29, 68, 100);
            Controls.Add(split); Controls.Add(status); Controls.Add(bar); Controls.Add(help);
            paragraphs.SelectedIndexChanged += (s, e) => { var p = paragraphs.SelectedItem as TemplateParagraph; text.Text = p == null ? "" : p.Text; };
            Shown += (s, e) => { if (fields.SelectedIndex < 0) fields.SelectedIndex = fields.FindStringExact("项目名称"); };
            RefreshList(null, 0);
        }
        protected override void OnLoad(EventArgs e) {
            base.OnLoad(e); float factor = DeviceDpi / 96F; if (factor != 1F) Scale(new SizeF(factor, factor));
            AutoScaleMode = AutoScaleMode.Dpi; AutoScaleDimensions = new SizeF(DeviceDpi, DeviceDpi);
            var area = Screen.FromControl(this).WorkingArea; Size = new Size(Math.Min(Width, area.Width), Math.Min(Height, area.Height));
        }
        void AddButton(Control bar, string title, Action action) { var b = new Button { Text = title, Width = 140, Height = 34 }; b.Click += (s, e) => action(); bar.Controls.Add(b); }
        void RefreshList(string part, int index) {
            paragraphs.Items.Clear(); var items = Core.TemplateParagraphs(Result); paragraphs.Items.AddRange(items.Cast<object>().ToArray());
            var selected = items.FirstOrDefault(x => x.Part == part && x.Index == index); if (selected != null) paragraphs.SelectedItem = selected; else if (items.Count > 0) paragraphs.SelectedIndex = 0;
        }
    }
}
