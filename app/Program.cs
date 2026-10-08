using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;

[assembly: System.Runtime.Versioning.TargetFramework(".NETFramework,Version=v4.8")]

namespace SettlementReview {
    public static partial class Core {
        public const string Version = "0.15.1";
        public static readonly string Root = AppDomain.CurrentDomain.BaseDirectory;
        public static readonly string Data = Path.Combine(Root, "data");
        public static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public static readonly Field[] Basic = {
            new Field("项目名称", "填写本次审核的工程全称"), new Field("报告编号", "完整填写本次编号；不再自动追加工程名称，所需后缀可自行填写"),
            new Field("审核单位"), new Field("委托单位"), new Field("建设单位"), new Field("代建单位", "无代建单位时填写“无”"),
            new Field("施工单位"), new Field("监理单位"), new Field("批复总投资", "写明数值和单位；无资料时填写待确认"),
            new Field("合同金额", "单位：元"), new Field("报告日期", "从日历中选择报告签发日期"),
            new Field("工程概况", "先显示模板内容，请按当前项目修改", 4)
        };
        public static readonly Field[] Schedule = {
            new Field("合同工期日历天数", "模板中的合同工期，可修改"), new Field("计划开工"), new Field("计划竣工"), new Field("实际开工"), new Field("实际竣工"),
            new Field("工期补充说明", "先沿用模板补充内容；顺延批准、罚款等请核对后修改", 4)
        };
        public static readonly Field[] Body = {
            new Field("审核依据及原则", "可手动编辑或从本次模板选择条目；请核对合同约定、版本和适用时间", 8),
            new Field("基期价格", "填写基期年月、造价信息或价格依据", 3),
            new Field("工程下浮率", "例如：12.08%；可补充适用说明", 2),
            new Field("工期延期报告", "填写延期原因、批准情况及处理结果；无则写“无”", 4),
            new Field("人员到位说明", "填写人员到位情况及处理依据；无则写“无”", 3),
            new Field("标化工地说明", "填写合同要求和费用处理情况；无则写“无”", 3),
            new Field("合同标内调整说明", "逐项写明项目、原因、依据和核增或核减金额；无则写“无”", 5),
            new Field("合同标外调整说明", "工程量与单价均变化时写明混合原因；资料不足时标待确认", 5),
            new Field("联系单调整说明", "填写联系单编号、核对情况、原因及金额；无则写“无”", 5),
            new Field("其他调整说明", "其他调整、取费与税金的影响；无则写“无”", 4),
            new Field("待确认及保留事项", "保留未解决事项及对结论的影响；无则写“无”", 4),
            new Field("附件清单", "逐行列出实际附件；软件不自动生成附件文件", 3)
        };
        public static readonly Field[] Amounts = {
            new Field("送审金额", "单位：元"), new Field("审定金额", "单位：元", 1, "审核金额"),
            new Field("核减合计", "单位：元；没有则填0", 1, "核减金额"), new Field("审增合计", "单位：元；没有则填0", 1, "核增金额"),
            new Field("净核减金额", "", 1, "净核减 / 净核增金额")
        };
        public static IEnumerable<Field> Fields { get { return Basic.Concat(Schedule).Concat(Body).Concat(Amounts); } }
        public static readonly string[] MatterFields = { "基期价格", "工程下浮率", "工期延期报告", "人员到位说明", "标化工地说明" };
        public static readonly string[] Derived = { "工期情况", "计划工期日历天数", "实际工期日历天数", "工期延期天数", "有关事项说明", "净核减率", "净额方向", "净额绝对金额", "金额汇总说明", "审核结论", "修正造价情况说明", "签发日期", "审核依据及原则", "审核内容", "审核程序及方法", "抄送单位", "审核金额", "核减金额", "核增金额", "审核人员", "复核人员", "税费口径", "金额比较口径", "审核结论补充说明" };
        public static HashSet<string> Allowed { get { return new HashSet<string>(Fields.Select(f => f.Name).Concat(Derived)); } }
        public static string Get(Project p, string key) { string v; return p.Values.TryGetValue(key, out v) ? v.Trim() : ""; }
        public static decimal Money(string value, string key) {
            decimal result;
            if (!Regex.IsMatch(value, @"^\d{1,15}(\.\d{1,2})?$" ) || !decimal.TryParse(value, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out result))
                throw new InvalidDataException(key + "：请填写非负金额，单位为元，最多15位整数和2位小数。不要输入逗号或“万元”。");
            return result;
        }
        public static string F(decimal d) { return d.ToString("0.00", CultureInfo.InvariantCulture); }
        public static Dictionary<string, string> Values(Project p, bool draft, IEnumerable<string> required = null) {
            return ReportValueBuilder.Build(p, draft, required);
        }
        // Compatibility facade for existing UI, template code and project files.
        public static void Save(string path, Project p) { ProjectStore.Save(path, p); }
        public static Project Load(string path) { return ProjectStore.Load(path); }
        public static void Atomic(string path, byte[] bytes) { ProjectStore.Atomic(path, bytes); }
        public static string SafeName(string name) {
            foreach (char c in Path.GetInvalidFileNameChars()) name = name.Replace(c, '_');
            name = name.Trim().TrimEnd('.'); if (name == "") name = "未命名项目";
            return name.Length > 60 ? name.Substring(0, 60) : name;
        }
        public static string ResolveTemplate(string path) { return Path.IsPathRooted(path) ? path : Path.Combine(Root, path); }
        public static void Open(string path) { System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true }); }

        static XElement P(string text, string style = "Normal") {
            return new XElement(W + "p", new XElement(W + "pPr", new XElement(W + "pStyle", new XAttribute(W + "val", style))),
                new XElement(W + "r", new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), text)));
        }
        static void WritePart(ZipArchive zip, string name, string xml) {
            var entry = zip.CreateEntry(name); using (var sw = new StreamWriter(entry.Open(), new UTF8Encoding(false))) sw.Write(xml);
        }
        public static byte[] FallbackBuiltIn() {
            var body = new XElement(W + "body");
            Action<string, string> add = (s, style) => body.Add(P(s, style));
            add("{{审核单位}}", "Company"); add("{{报告编号}}", "Right");
            add("关于{{项目名称}}的", "Title"); add("结算审核报告", "Title");
            add("{{委托单位}}：", "Normal");
            add("我公司接受贵公司的委托，对{{项目名称}}的结算进行了审核。委托方的责任是提供真实、合法、完整的工程资料，我们的责任是按照客观、公正、合理和实事求是的原则对该工程的结算进行审核。现将审核情况报告如下：", "Normal");
            add("一 项目概况", "Heading1");
            foreach (string key in new[] { "批复总投资", "建设单位", "代建单位", "施工单位", "监理单位" }) add(key + "：{{" + key + "}}", "Normal");
            add("合同金额：{{合同金额}} 元。", "Normal"); add("{{工程概况}}", "Normal"); add("{{工期情况}}", "Normal");
            add("二 审核依据及原则", "Heading1"); add("{{审核依据及原则}}", "Normal");
            add("三 审核内容", "Heading1"); add("{{审核内容}}", "Normal");
            add("四 审核程序及方法", "Heading1"); add("{{审核程序及方法}}", "Normal");
            add("五 有关事项说明", "Heading1"); add("{{有关事项说明}}", "Normal");
            add("六 修正造价情况说明", "Heading1"); add("{{金额汇总说明}}", "Normal");
            add("合同标内工程量及相关调整", "Heading2"); add("{{合同标内调整说明}}", "Normal");
            add("合同标外工程量及综合单价调整", "Heading2"); add("{{合同标外调整说明}}", "Normal");
            add("联系单工程量及综合单价调整", "Heading2"); add("{{联系单调整说明}}", "Normal");
            add("其他调整", "Heading2"); add("{{其他调整说明}}", "Normal");
            add("七 审核结论", "Heading1"); add("{{审核结论}}", "Normal");
            add("待确认及保留事项：{{待确认及保留事项}}", "Normal");
            add("八 附件", "Heading1"); add("{{附件清单}}", "Normal");
            add("{{审核单位}}", "Right"); add("{{报告日期}}", "Right");
            add("主题词：工程 结算 审核", "Small"); add("抄送：{{抄送单位}}", "Small");
            body.Add(new XElement(W + "sectPr", new XElement(W + "pgSz", new XAttribute(W + "w", "11906"), new XAttribute(W + "h", "16838")),
                new XElement(W + "pgMar", new XAttribute(W + "top", "1440"), new XAttribute(W + "right", "1440"), new XAttribute(W + "bottom", "1440"), new XAttribute(W + "left", "1701"), new XAttribute(W + "header", "720"), new XAttribute(W + "footer", "720"), new XAttribute(W + "gutter", "0"))));
            var doc = new XDocument(new XElement(W + "document", new XAttribute(XNamespace.Xmlns + "w", W), body));
            string styles = "<w:styles xmlns:w='" + W + "'><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii='Times New Roman' w:eastAsia='宋体' w:hAnsi='Times New Roman'/><w:sz w:val='24'/><w:color w:val='000000'/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after='120' w:line='360' w:lineRule='auto'/><w:widowControl/></w:pPr></w:pPrDefault></w:docDefaults>";
            styles += "<w:style w:type='paragraph' w:default='1' w:styleId='Normal'><w:name w:val='Normal'/></w:style>";
            foreach (string name in new[] { "Company", "Title", "Heading1", "Heading2", "Right", "Small" }) {
                string align = name == "Company" || name == "Title" ? "center" : name == "Right" ? "right" : "left";
                string size = name == "Company" ? "36" : name == "Title" ? "32" : name == "Small" ? "20" : "24";
                styles += "<w:style w:type='paragraph' w:styleId='" + name + "'><w:name w:val='" + (name == "Heading1" ? "heading 1" : name == "Heading2" ? "heading 2" : name) + "'/><w:basedOn w:val='Normal'/><w:pPr><w:jc w:val='" + align + "'/>" + (name.StartsWith("Heading") || name == "Title" || name == "Company" ? "<w:keepNext/><w:spacing w:before='200' w:after='160'/>" : "") + (name.StartsWith("Heading") ? "<w:outlineLvl w:val='" + (name == "Heading1" ? "0" : "1") + "'/>" : "") + "</w:pPr><w:rPr><w:sz w:val='" + size + "'/>" + (name.StartsWith("Heading") || name == "Company" || name == "Title" ? "<w:b/><w:rFonts w:eastAsia='黑体'/>" : "") + "</w:rPr></w:style>";
            }
            styles += "</w:styles>";
            using (var ms = new MemoryStream()) {
                using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, true)) {
                    WritePart(zip, "[Content_Types].xml", "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Default Extension='xml' ContentType='application/xml'/><Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/><Override PartName='/word/styles.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml'/></Types>");
                    WritePart(zip, "_rels/.rels", "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/></Relationships>");
                    WritePart(zip, "word/_rels/document.xml.rels", "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles' Target='styles.xml'/></Relationships>");
                    WritePart(zip, "word/document.xml", doc.ToString(SaveOptions.DisableFormatting)); WritePart(zip, "word/styles.xml", styles);
                } return ms.ToArray();
            }
        }
        static XDocument Xml(ZipArchiveEntry entry) {
            using (var sr = System.Xml.XmlReader.Create(entry.Open(), new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 32000000 })) return XDocument.Load(sr, LoadOptions.PreserveWhitespace);
        }
        static bool IsTextPart(string name) { return name.StartsWith("word/", StringComparison.Ordinal) && name.EndsWith(".xml", StringComparison.Ordinal) && !name.Contains("/theme/"); }
        public static HashSet<string> ValidateTemplate(byte[] bytes) {
            var fields = new HashSet<string>();
            using (var ms = new MemoryStream(bytes)) using (var zip = new ZipArchive(ms, ZipArchiveMode.Read)) {
                if (zip.GetEntry("word/document.xml") == null || zip.GetEntry("[Content_Types].xml") == null) throw new InvalidDataException("不是有效的.docx文件。旧式.doc请先在WPS或Word中另存为.docx。");
                if (zip.Entries.Sum(e => e.Length) > 32000000 || zip.Entries.Count > 3000) throw new InvalidDataException("模板过大，请使用32MB以内的模板。");
                foreach (var entry in zip.Entries) {
                    if (entry.FullName.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("请选择不含宏的.docx模板。");
                    if (!IsTextPart(entry.FullName)) continue;
                    var xml = Xml(entry);
                    if (xml.Descendants(W + "documentProtection").Any() || xml.Descendants(W + "ins").Any() || xml.Descendants(W + "del").Any()) throw new InvalidDataException("模板副本预处理未完成，请重新导入。");
                    foreach (var p in xml.Descendants(W + "p")) {
                        string t = string.Concat(p.Descendants(W + "t").Where(x => x.Ancestors(W + "p").First() == p).Select(x => x.Value));
                        var matches = Regex.Matches(t, @"\{\{([^{}]+)\}\}");
                        foreach (Match m in matches) fields.Add(m.Groups[1].Value);
                        string rest = Regex.Replace(t, @"\{\{[^{}]+\}\}", "");
                        if (rest.Contains("{{") || rest.Contains("}}")) throw new InvalidDataException("模板有未闭合或跨段落的字段：" + t);
                    }
                    foreach (var t in xml.Descendants().Where(x => (x.Name.LocalName == "instrText" || x.Name.LocalName == "delText") && x.Value.Contains("{{"))) throw new InvalidDataException("字段位于不支持的位置，请移到普通正文、表格或页眉页脚文本。");
                    if (xml.Descendants().Where(x => !x.HasElements && (x.Value.Contains("{{") || x.Value.Contains("}}"))).Any(x => x.Name != W + "t" || !x.Ancestors(W + "p").Any()) || xml.Descendants().Attributes().Any(a => a.Value.Contains("{{") || a.Value.Contains("}}")))
                        throw new InvalidDataException("模板含普通段落以外的字段标记。请将字段放入正文、表格、页眉页脚或文本框的普通文本中。");
                }
            }
            if (fields.Count == 0) throw new InvalidDataException("模板没有{{字段名}}标记。请先导出字段模板作为参考，将标记放入您自己的.docx副本后导入。");
            var unknown = fields.Where(f => !Allowed.Contains(f)).ToArray();
            if (unknown.Length > 0) throw new InvalidDataException("模板含未识别字段：" + string.Join("、", unknown) + "。可用字段见“模板与导出”。");
            return fields;
        }
        static void ReplaceText(XDocument xml, Dictionary<string, string> values) {
            foreach (var p in xml.Descendants(W + "p").ToList()) {
                var nodes = p.Descendants(W + "t").Where(x => x.Ancestors(W + "p").First() == p).ToList();
                string full = string.Concat(nodes.Select(n => n.Value));
                foreach (Match match in Regex.Matches(full, @"\{\{([^{}]+)\}\}").Cast<Match>().Reverse()) {
                    string value; if (!values.TryGetValue(match.Groups[1].Value, out value)) throw new InvalidDataException("未知字段：" + match.Value);
                    int start = match.Index, end = start + match.Length, cursor = 0;
                    foreach (var n in nodes) {
                        int next = cursor + n.Value.Length;
                        if (cursor < end && next > start) {
                            int a = Math.Max(0, start - cursor), b = Math.Min(n.Value.Length, end - cursor);
                            n.Value = n.Value.Substring(0, a) + (cursor <= start ? value : "") + n.Value.Substring(b);
                            n.SetAttributeValue(XNamespace.Xml + "space", "preserve");
                        } cursor = next;
                    }
                }
                foreach (var n in nodes.Where(n => n.Value.Contains("\n") || n.Value.Contains("\r")).ToList()) {
                    string[] lines = n.Value.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
                    var parts = new List<XElement>();
                    for (int i = 0; i < lines.Length; i++) { if (i > 0) parts.Add(new XElement(W + "br")); parts.Add(new XElement(W + "t", new XAttribute(XNamespace.Xml + "space", "preserve"), lines[i])); }
                    n.ReplaceWith(parts);
                }
            }
        }
        public static byte[] Export(Project project, byte[] template, bool draft) {
            template = PrepareTemplate(template); var templateFields = ValidateTemplate(template); var values = Values(project, draft, RequiredForTemplate(templateFields));
            if (!draft && project.Excel != null && project.Excel.AutoLink && !templateFields.Contains("修正造价情况说明")) {
                var missing = ExcelReview.Targets.Where(f => values[f] != "无" && !templateFields.Contains(f)).ToArray();
                if (missing.Length > 0) throw new InvalidDataException("当前模板未关联以下 Excel 正文位置：" + string.Join("、", missing) + "。请在第一步关联填写位置后再导出。");
            }
            template = FormatReviewTemplate(project, template, values);
            template = RemoveAbsentSections(template, project, values);
            using (var input = new MemoryStream(template)) using (var zin = new ZipArchive(input, ZipArchiveMode.Read)) using (var output = new MemoryStream()) {
                using (var zout = new ZipArchive(output, ZipArchiveMode.Create, true)) {
                    foreach (var part in zin.Entries) {
                        var target = zout.CreateEntry(part.FullName);
                        using (var dest = target.Open()) {
                            if (IsTextPart(part.FullName)) {
                                var xml = Xml(part); RemoveAutomaticNumberSuffix(xml);
                                bool hasField = xml.Descendants(W + "p").Any(p => string.Concat(p.Descendants(W + "t").Select(t => t.Value)).Contains("{{"));
                                if (hasField || (draft && part.FullName == "word/document.xml")) {
                                    RemoveAbsentProjectLines(xml, project);
                                    ExpandBasisParagraphs(xml, values);
                                    ReplaceText(xml, values);
                                    if (draft && part.FullName == "word/document.xml") xml.Root.Element(W + "body").AddFirst(P("草稿  待审核确认", "Normal"));
                                    using (var writer = new StreamWriter(dest, new UTF8Encoding(false))) xml.Save(writer, SaveOptions.DisableFormatting);
                                } else using (var source = part.Open()) source.CopyTo(dest);
                            } else using (var source = part.Open()) source.CopyTo(dest);
                        }
                    }
                } return output.ToArray();
            }
        }
        public static string Preview(Project p, byte[] template = null) {
            var v = Values(p, true);
            using (var ms = new MemoryStream(Export(p, template ?? BuiltIn(), true))) using (var zip = new ZipArchive(ms, ZipArchiveMode.Read)) {
                var xml = Xml(zip.GetEntry("word/document.xml"));
                var ids = ReviewNumberIds(zip); var counts = ids.ToDictionary(id => id, id => 0);
                return string.Join("\r\n\r\n", xml.Descendants(W + "p").Select(par => { var num = par.Element(W + "pPr") == null ? null : par.Element(W + "pPr").Element(W + "numPr"); string id = num == null || num.Element(W + "numId") == null ? null : (string)num.Element(W + "numId").Attribute(W + "val"); string prefix = id != null && ids.Contains(id) ? (++counts[id]) + ". " : ""; return prefix + string.Concat(par.Descendants().Where(x => x.Name == W + "t" || x.Name == W + "br").Select(x => x.Name == W + "br" ? "\r\n" : x.Value)); }));
            }
        }
    }

    public partial class MainForm : Form {
        readonly bool testMode;
        Project current = new Project(); bool loading, dirty; string selectedTemplate = "";
        readonly Dictionary<string, TextBox> inputs = new Dictionary<string, TextBox>();
        readonly ListBox projects = new ListBox(); readonly TextBox search = new TextBox(); readonly TabControl tabs = new TabControl();
        readonly Label status = new Label(), amountStatus = new Label(), templateStatus = new Label();
        readonly RichTextBox preview = new RichTextBox(); readonly Panel wordPreviewHost = new Panel();
        readonly TextBox legacyNotes = new TextBox();
        readonly Dictionary<string, DateTimePicker> dateInputs = new Dictionary<string, DateTimePicker>();
        readonly TextBox schedulePreview = new TextBox(); readonly CheckBox inclusiveDays = new CheckBox(); readonly Label templateHomeStatus = new Label();
        readonly CheckBox netAmountAuto = new CheckBox(); bool netAmountUpdating;
        readonly Timer autosave = new Timer(); readonly string projectData, projectDir, draftPath, uiPreferencesPath; readonly bool isolatedData;
        List<Project> catalog = new List<Project>();
        public MainForm(bool test = false, string testDataRoot = null) {
            testMode = test; isolatedData = test && testDataRoot != null; projectData = isolatedData ? Path.GetFullPath(testDataRoot) : Core.Data; projectDir = Path.Combine(projectData, "projects"); draftPath = Path.Combine(projectData, "draft.json"); uiPreferencesPath = Path.Combine(projectData, "ui-preferences.json");
            Text = "结算审核工作台  ·  试用版 " + Core.Version; Font = new Font("Microsoft YaHei UI", 10.5F, FontStyle.Regular, GraphicsUnit.Point); ForeColor = Color.FromArgb(30, 45, 61); BackColor = Color.FromArgb(245, 247, 250);
            AutoScaleMode = AutoScaleMode.None;
            Size = new Size(1360, 920); MinimumSize = new Size(1060, 720); StartPosition = FormStartPosition.CenterScreen;
            var top = new Panel { Dock = DockStyle.Top, Height = 48, BackColor = Color.FromArgb(24, 49, 70), Padding = new Padding(16, 7, 16, 5) };
            top.Controls.Add(new Label { Text = "结算审核工作台", Font = new Font(Font.FontFamily, 14, FontStyle.Bold), ForeColor = Color.White, Dock = DockStyle.Top, Height = 30 });
            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 40, Padding = new Padding(16, 8, 12, 6), BackColor = Color.White }; status.Dock = DockStyle.Fill; bottom.Controls.Add(status);
            var left = new Panel { Dock = DockStyle.Left, Width = 235, Padding = new Padding(14), BackColor = Color.FromArgb(232, 238, 244) };
            var leftTitle = new Label { Text = "本地项目", Dock = DockStyle.Top, Height = 30, Font = new Font(Font, FontStyle.Bold) };
            search.Dock = DockStyle.Top; search.Margin = new Padding(0, 0, 0, 10); search.TextChanged += (s, e) => Filter();
            projects.Dock = DockStyle.Fill; projects.BorderStyle = BorderStyle.None; projects.HorizontalScrollbar = true; projects.BackColor = left.BackColor; projects.IntegralHeight = false; projects.ItemHeight = 30; projects.DoubleClick += (s, e) => OpenSelected();
            var leftActions = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 176, FlowDirection = FlowDirection.TopDown };
            leftActions.Controls.Add(Button("新建项目", (s, e) => NewProject(), 192));
            leftActions.Controls.Add(Button("打开项目文件…", (s, e) => OpenProjectFile(), 192));
            leftActions.Controls.Add(Button("保存项目到…", (s, e) => SaveProjectAs(), 192));
            leftActions.Controls.Add(Button("查看审核版本", (s, e) => ShowProjectVersions(), 192));

            var leftHelp = new Label { Dock = DockStyle.Bottom, Height = 52, Text = "上方输入关键字筛选\n双击项目继续编辑", ForeColor = Color.DimGray };
            left.Controls.Add(projects); left.Controls.Add(leftHelp); left.Controls.Add(leftActions); left.Controls.Add(search); left.Controls.Add(leftTitle);
            var right = new Panel { Dock = DockStyle.Fill, Padding = new Padding(20, 12, 20, 12) };
            var actions = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 42 };

            actions.Controls.Add(Button("保存审核副本", (s, e) => CopyProject(), 150));
            actions.Controls.Add(Button("打开数据目录", (s, e) => Run(() => { Directory.CreateDirectory(Core.Data); Core.Open(Core.Data); }), 150));
            var currentVersion = new Label { Text = "本地保存  ·  模板与 Excel 协同审核", AutoSize = true, Padding = new Padding(10, 9, 0, 0), ForeColor = Color.DimGray }; actions.Controls.Add(currentVersion);
            tabs.Dock = DockStyle.Fill; tabs.Padding = new Point(10, 12); tabs.DrawMode = TabDrawMode.OwnerDrawFixed;
            tabs.DrawItem += (s,e) => { bool active = e.Index == tabs.SelectedIndex; var rect = e.Bounds; using (var brush = new SolidBrush(active ? Color.White : Color.FromArgb(237, 242, 247))) e.Graphics.FillRectangle(brush, rect); TextRenderer.DrawText(e.Graphics, tabs.TabPages[e.Index].Text, Font, rect, active ? Color.FromArgb(20, 79, 122) : Color.FromArgb(78, 96, 112), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter); if (active) using (var pen = new Pen(Color.FromArgb(39, 117, 168), 3)) e.Graphics.DrawLine(pen, rect.Left + 10, rect.Bottom - 2, rect.Right - 10, rect.Bottom - 2); };
            BuildTemplateHome();
            var basicTab = AddFields("2  项目信息", Core.Basic); var scheduleTab = AddFields("3  工期情况", Core.Schedule); var bodyTab = AddFields("4  审核正文", Core.Body);
            BuildExcelTab(); var moneyTab = AddFields("6  金额结论", Core.Amounts);
            inclusiveDays.Text = "首尾均计（开工、竣工当天都计入，与当前模板一致）"; inclusiveDays.Checked = true; inclusiveDays.Dock = DockStyle.Top; inclusiveDays.Height = 40; inclusiveDays.CheckedChanged += (s, e) => Changed(); scheduleTab.Controls.Add(inclusiveDays);
            schedulePreview.Multiline = true; schedulePreview.ReadOnly = true; schedulePreview.ScrollBars = ScrollBars.Vertical; schedulePreview.BackColor = Color.FromArgb(231, 240, 247); schedulePreview.Dock = DockStyle.Bottom; schedulePreview.Height = 150; scheduleTab.Controls.Add(schedulePreview);
            Resize += (s, e) => { float scale = DeviceDpi / 96F; schedulePreview.Height = (int)((ClientSize.Height / scale < 760 ? 100 : 150) * scale); };
            basicTab.Controls.Add(new Label { Dock = DockStyle.Top, Height = 42, Padding = new Padding(16, 10, 8, 0), Text = "抄送单位自动按建设单位、施工单位、代建单位生成，重复单位只列一次。", ForeColor = Color.FromArgb(29, 68, 100) });
            bodyTab.Controls.Add(new Label { Dock = DockStyle.Top, Height = 64, Padding = new Padding(16, 10, 8, 0), Text = "审核依据及原则可编辑或选择；审核内容、审核程序及方法沿用模板。\n合同标内起的四项调整说明与 Excel 联动；在 Excel 中修改原因、报告归属和采用项。", ForeColor = Color.FromArgb(29, 68, 100) });
            legacyNotes.Dock = DockStyle.Top; legacyNotes.Height = 100; legacyNotes.Multiline = true; legacyNotes.ReadOnly = true; legacyNotes.ScrollBars = ScrollBars.Vertical; legacyNotes.Visible = false; bodyTab.Controls.Add(legacyNotes);
            amountStatus.AutoSize = false; amountStatus.Dock = DockStyle.Bottom; amountStatus.Height = 160; amountStatus.Padding = new Padding(16, 10, 12, 8); amountStatus.BackColor = Color.FromArgb(231, 240, 247); moneyTab.Controls.Add(amountStatus);
            BuildExportTab(); right.Controls.Add(tabs); right.Controls.Add(actions);
            Controls.Add(right); Controls.Add(left); Controls.Add(bottom); Controls.Add(top);
            autosave.Interval = 1500; autosave.Tick += (s, e) => { autosave.Stop(); if (!loading && !testMode) Run(() => { Gather(); Core.Save(draftPath, current); status.Text = "草稿已保存  " + DateTime.Now.ToString("HH:mm:ss") + "  ·  点击左上角“保存项目到…”选择文件位置"; }); };
            FormClosing += (s, e) => { if (test) return; try { Gather(); Core.Save(draftPath, current); } catch (Exception ex) { if (MessageBox.Show("草稿保存失败：" + ex.Message + "\n仍要关闭吗？", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) e.Cancel = true; } };
            if (!test) {
                Directory.CreateDirectory(projectDir); RefreshProjects();
                if (File.Exists(draftPath)) Run(() => { current = Core.Load(draftPath); });
            }
            LoadIntoForm(current); status.Text = "请先导入模板或使用默认模板，再核对回填内容。";
        }
        protected override void OnLoad(EventArgs e) {
            base.OnLoad(e);
            float factor = DeviceDpi / 96F;
            if (factor != 1F) Scale(new SizeF(factor, factor));
            foreach (DataGridViewColumn column in excelGrid.Columns) column.Width = (int)(column.Width * factor);
            excelGrid.ColumnHeadersHeight = (int)(40 * factor);
            foreach (DataGridViewRow row in excelGrid.Rows) row.Height = (int)(36 * factor);
            AutoScaleMode = AutoScaleMode.Dpi; AutoScaleDimensions = new SizeF(DeviceDpi, DeviceDpi);
            var area = Screen.FromControl(this).WorkingArea;
            MinimumSize = new Size(Math.Min(MinimumSize.Width, area.Width), Math.Min(MinimumSize.Height, area.Height));
            Size = new Size(Math.Min(Width, area.Width), Math.Min(Height, area.Height));
            excelSplitInitialized = false; AdjustExcelLayout(); schedulePreview.Height = (int)((ClientSize.Height / factor < 760 ? 100 : 150) * factor);
        }
        Button Button(string text, EventHandler action, int width = 140) {
            var b = new Button { Text = text, Width = width, Height = 36, FlatStyle = FlatStyle.Flat, BackColor = Color.White, Margin = new Padding(0, 0, 8, 7), Cursor = Cursors.Hand };
            b.FlatAppearance.BorderColor = Color.FromArgb(188, 203, 215); b.Click += action; return b;
        }
        void Run(Action action) { try { action(); } catch (Exception ex) { status.Text = "操作未完成，数据未清除。"; MessageBox.Show(ex.Message, "请检查", MessageBoxButtons.OK, MessageBoxIcon.Warning); } }
        TabPage AddFields(string title, IEnumerable<Field> fields) {
            var page = new TabPage(title) { BackColor = Color.White, Padding = new Padding(8) };
            var scroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true };
            var grid = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2, Padding = new Padding(10, 10, 22, 20) };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 160)); grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            int row = 0;
            foreach (var field in fields) {
                grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                var label = new Label { Text = field.Label, AutoSize = true, Margin = new Padding(0, 10, 12, 12), MaximumSize = new Size(155, 0) };
                if (Core.DateInputFields.Contains(field.Name)) {
                    var picker = new DateTimePicker { Dock = DockStyle.Top, Format = DateTimePickerFormat.Custom, CustomFormat = " ", ShowCheckBox = true, Checked = false, MinDate = new DateTime(1900, 1, 1), MaxDate = new DateTime(2200, 12, 31) };
                    picker.ValueChanged += (s, e) => { picker.CustomFormat = picker.Checked ? "yyyy-MM-dd" : " "; Changed(); };
                    dateInputs[field.Name] = picker; var datePanel = new Panel { Dock = DockStyle.Fill, Height = 64, Margin = new Padding(0, 4, 0, 8) }; datePanel.Controls.Add(picker);
                    grid.Controls.Add(label, 0, row); grid.Controls.Add(datePanel, 1, row++); continue;
                }
                var box = new TextBox { Multiline = field.Lines > 1, AcceptsReturn = field.Lines > 1, ScrollBars = field.Lines > 1 ? ScrollBars.Vertical : ScrollBars.None, Dock = DockStyle.Top, Height = field.Lines > 1 ? (field.Name == "审核依据及原则" ? 150 : Math.Min(field.Lines, 3) * 24) : 30, MaxLength = 20000 };
                box.TextChanged += (s, e) => { if (field.Name == "工程下浮率" && !loading && !discountUpdating) { discountUpdating = true; discountAuto.Checked = false; discountUpdating = false; } if (field.Name == "净核减金额" && !loading && !netAmountUpdating) { netAmountUpdating = true; netAmountAuto.Checked = false; netAmountUpdating = false; } Changed(); }; inputs[field.Name] = box;
                var cell = new Panel { Dock = DockStyle.Fill, Height = box.Height + 34, Margin = new Padding(0, 4, 0, 8) };
                if (field.Hint != "") cell.Controls.Add(new Label { Text = ExcelReview.Targets.Contains(field.Name) ? "导入 Excel 后自动联动；关闭联动后可手动编辑当前正文" : field.Hint, ForeColor = Color.DimGray, Dock = DockStyle.Bottom, Height = 26, Padding = new Padding(0, 3, 0, 0) });
                cell.Controls.Add(box); if (field.Name == "审核依据及原则") BuildBasisControls(cell, box); if (field.Name == "工程下浮率") BuildDiscountInputs(cell, box); if (field.Name == "净核减金额") { cell.Height += 28; netAmountAuto.Text = "自动按“送审金额 − 审核金额”生成；取消勾选后可手动修改"; netAmountAuto.Dock = DockStyle.Bottom; netAmountAuto.Height = 28; netAmountAuto.CheckedChanged += (s,e) => { if (!loading && !netAmountUpdating) Changed(); }; cell.Controls.Add(netAmountAuto); } grid.Controls.Add(label, 0, row); grid.Controls.Add(cell, 1, row++);
            }
            scroll.Controls.Add(grid); page.Controls.Add(scroll); tabs.TabPages.Add(page); return page;
        }
        void BuildTemplateHome() {
            var page = new TabPage("1  导入模板") { BackColor = Color.White, Padding = new Padding(24) };
            var content = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false };
            content.Controls.Add(new Label { Text = "先选择本次报告模板", Width = 760, Height = 50, Font = new Font(Font.FontFamily, 18, FontStyle.Bold) });
            content.Controls.Add(new Label { Text = "导入后按模板刷新项目信息、工期及有关事项，原项目会先自动保存一份副本。\n模板中的项目名称、编号和日期也会显示，请按本次工程核对；审核金额保留您的输入。", Width = 760, Height = 75 });
            var buttons = new FlowLayoutPanel { Width = 760, Height = 52 };
            buttons.Controls.Add(Button("导入 Word 模板", (s, e) => ImportTemplate(), 180));
            buttons.Controls.Add(Button("使用默认模板并回填", (s, e) => UseDefaultTemplate(), 200));
            buttons.Controls.Add(Button("设置填写位置", (s, e) => EditTemplateMappings(), 160));
            content.Controls.Add(buttons); templateHomeStatus.Width = 760; templateHomeStatus.Height = 90; templateHomeStatus.ForeColor = Color.FromArgb(29, 68, 100); content.Controls.Add(templateHomeStatus);
            content.Controls.Add(new Label { Text = "后续步骤：项目信息 → 工期情况 → 审核正文 → Excel对比 → 金额结论 → 报告导出\n\nExcel 对比可预览送审、审查、核增、核减明细，选择调整原因后写入当前报告。\n工期自动计算日历天数，金额页实时核对净核减金额。", Width = 810, Height = 140 });
            page.Controls.Add(content); tabs.TabPages.Add(page);
        }
        void BackupBeforeTemplate() { if (!testMode && current.Values.Any(p => !string.IsNullOrWhiteSpace(p.Value))) Core.Save(Path.Combine(Core.Data, "template-backups", DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + "_" + current.Id + ".json"), current); }
        void ApplyTemplateToForm(byte[] source, byte[] prepared, string path) { Gather(); BackupBeforeTemplate(); current.Template = path; Core.ApplyTemplateDefaults(current, source, prepared, true); if (current.Excel != null) current.Excel.WordingMode = ExcelReview.DetectWordingMode(source); foreach (string key in Core.DiscountInputs) current.Values[key] = ""; current.Values["下浮率模式"] = "自动"; LoadIntoForm(current); Changed(); tabs.SelectedIndex = 1; }
        void UseDefaultTemplate() { Run(() => { ApplyTemplateToForm(Core.SourceTemplate(), Core.BuiltIn(), ""); status.Text = "已按默认模板刷新项目信息和工期，审核金额保留，请核对本次项目资料。"; }); }
        void UpdateSchedule() {
            try { Gather(); var schedule = Core.ScheduleValues(current, true); schedulePreview.ForeColor = Color.FromArgb(27, 78, 103); schedulePreview.Text = "计划工期：" + schedule["计划工期日历天数"] + " 日历天    实际工期：" + schedule["实际工期日历天数"] + " 日历天\r\n\r\n" + schedule["工期情况"]; }
            catch (Exception ex) { schedulePreview.ForeColor = Color.Firebrick; schedulePreview.Text = ex.Message; }
        }
        void BuildExportTab() {
            var page = new TabPage("7  报告导出") { BackColor = Color.White, Padding = new Padding(16) };
            var panel = new Panel { Dock = DockStyle.Top, Height = 161 };
            templateStatus.Dock = DockStyle.Top; templateStatus.Height = 60; templateStatus.ForeColor = Color.FromArgb(29, 68, 100);
            var fieldHint = new Label { Dock = DockStyle.Top, Height = 46, Text = "导入 .docx 后，以该模板为准。下方按最终 Word 结构显示字体、字号、加粗、下划线、对齐、缩进和编号。\n预览使用临时副本；复杂绘图、页眉页脚和精确分页请在 Word 或 WPS 中复核。", ForeColor = Color.DimGray };
            var exports = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 48 };
            exports.Controls.Add(Button("刷新 Word 预览", (s, e) => RefreshPreview(), 150));
            exports.Controls.Add(Button("导出草稿 Word", (s, e) => ExportReport(true), 150));
            var export = Button("生成审核稿 Word", (s, e) => ExportReport(false), 170); export.BackColor = Color.FromArgb(31, 85, 119); export.ForeColor = Color.White; exports.Controls.Add(export);
            exports.Controls.Add(Button("查看可用字段", (s, e) => MessageBox.Show(string.Join("  ", Core.Allowed.Select(f => "{{" + f + "}}")), "可用模板字段"), 140));
            panel.Controls.Add(exports); panel.Controls.Add(fieldHint); panel.Controls.Add(templateStatus);
            wordPreviewHost.Dock = DockStyle.Fill; wordPreviewHost.AutoScroll = true; wordPreviewHost.BackColor = Color.FromArgb(224, 228, 232); wordPreviewHost.Padding = new Padding(20, 14, 20, 14);
            preview.ReadOnly = true; preview.WordWrap = true; preview.DetectUrls = false; preview.ScrollBars = RichTextBoxScrollBars.Vertical; preview.BackColor = Color.White; preview.BorderStyle = BorderStyle.FixedSingle; preview.Font = new Font("SimSun", 12F); preview.ZoomFactor = 1F;
            wordPreviewHost.Controls.Add(preview); wordPreviewHost.Resize += (s,e) => LayoutWordPreview();
            page.Controls.Add(wordPreviewHost); page.Controls.Add(panel); tabs.TabPages.Add(page);
            tabs.SelectedIndexChanged += (s, e) => { if (tabs.SelectedTab == page) RefreshPreview(); };
        }
        void Gather() { foreach (var item in inputs) current.Values[item.Key] = item.Value.Text; foreach (var item in dateInputs) current.Values[item.Key] = item.Value.Checked ? item.Value.Value.ToString("yyyy-MM-dd") : ""; current.Values["工期口径"] = inclusiveDays.Checked ? "首尾均计" : "日期相减"; current.Values["下浮率模式"] = discountAuto.Checked ? "自动" : "手动"; current.Values["净额模式"] = netAmountAuto.Checked ? "自动" : "手动"; current.Template = selectedTemplate; }
        void Changed() { if (loading) return; Gather(); UpdateNetAmount(); Gather(); UpdateDiscount(); UpdateLinkedBody(); dirty = true; autosave.Stop(); if (!testMode) autosave.Start(); UpdateAmounts(); UpdateSchedule(); }
        void UpdateNetAmount() {
            if (!netAmountAuto.Checked || netAmountUpdating || !inputs.ContainsKey("净核减金额")) return;
            try {
                string before = inputs["送审金额"].Text.Trim(), after = inputs["审定金额"].Text.Trim(); if (before == "" || after == "") return;
                decimal value = Core.Money(before, "送审金额") - Core.Money(after, "审核金额"); string text = (value < 0 ? "-" : "") + Core.F(Math.Abs(value));
                netAmountUpdating = true; bool wasLoading = loading; loading = true; try { inputs["净核减金额"].Text = text; current.Values["净核减金额"] = text; } finally { loading = wasLoading; netAmountUpdating = false; }
            } catch { }
        }
        void UpdateAmounts() {
            try { Gather(); string check = Core.AmountCheck(current); amountStatus.ForeColor = check.Contains("提示：") ? Color.Firebrick : Color.FromArgb(27, 78, 103); amountStatus.Text = check + "\n\n按当前模板生成：\n" + Core.ConclusionPreview(current, CurrentTemplate()); }
            catch (Exception ex) { amountStatus.ForeColor = Color.Firebrick; amountStatus.Text = ex.Message; }
        }
        void UpdateTemplate() { templateStatus.Text = selectedTemplate == "" ? "当前：您的《结算审核模板.docx》\n沿用原文档的样式、页眉页脚与三段固定审核正文。" : "当前：" + Path.GetFileName(selectedTemplate) + "\n模板已就绪；原文件保留，后续操作均使用副本。"; templateHomeStatus.Text = templateStatus.Text; }
        void LoadIntoForm(Project p) {
            projectFilePath = ""; loading = true; autosave.Stop(); current = p; selectedTemplate = p.Template ?? "";
            if (Core.Get(p, "模板审核依据") == "") p.Values["模板审核依据"] = Core.TemplateBasis(CurrentWordingTemplate());
            if (Core.Get(p, "审核依据及原则") == "") p.Values["审核依据及原则"] = Core.ProjectBasis(p);
            foreach (var item in inputs) item.Value.Text = Core.Get(p, item.Key);
            foreach (var item in dateInputs) { var date = Core.ParseDate(Core.Get(p, item.Key)); if (date.HasValue) item.Value.Value = date.Value; item.Value.Checked = date.HasValue; item.Value.CustomFormat = date.HasValue ? "yyyy-MM-dd" : " "; }
            inclusiveDays.Checked = Core.Get(p, "工期口径") != "日期相减";
            discountAuto.Checked = Core.Get(p, "下浮率模式") != "手动";
            netAmountAuto.Checked = Core.Get(p, "净额模式") == "自动" || Core.Get(p, "净额模式") == "" && Core.Get(p, "净核减金额") == "";
            UpdateDiscount();
            LoadExcelState();
            UpdateLinkedBody();
            string oldNotes = Core.Get(p, "有关事项说明"); legacyNotes.Visible = oldNotes != "" && oldNotes != "无";
            legacyNotes.Text = "旧版有关事项说明（保留供拆分填写；不会自动归类）：\r\n" + oldNotes;
            loading = false; UpdateNetAmount(); Gather(); dirty = false; UpdateAmounts(); UpdateTemplate(); UpdateSchedule();
        }
        void SaveProject() { if (testMode && !isolatedData) throw new InvalidOperationException("测试必须使用独立数据目录。"); Gather(); Core.Save(Path.Combine(projectDir, current.Id + ".json"), current); Core.Save(draftPath, current); dirty = false; RefreshProjects(); status.Text = "项目已保存  ·  " + DateTime.Now.ToString("HH:mm:ss"); }
        bool LeaveCurrent() {
            if (!dirty) return true;
            try { SaveProject(); return true; } catch (Exception ex) { MessageBox.Show("当前项目保存失败，暂不切换：" + ex.Message); return false; }
        }
        void NewProject() { if (!LeaveCurrent()) return; LoadIntoForm(new Project()); tabs.SelectedIndex = 0; status.Text = "已新建空白项目"; }
        void CopyProject() { Run(() => { SaveProject(); string path = Path.Combine(projectData, "copies", current.Id, DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + "_" + Guid.NewGuid().ToString("N").Substring(0, 6) + ".json"); Core.Save(path, current); status.Text = "审核副本已保存到当前工程，可点击左侧“查看审核版本”打开。"; }); }
        void OpenSelected() { var group = projects.SelectedItem as ProjectGroup; if (group == null || !LeaveCurrent()) return; var p = group.Latest; Run(() => { LoadIntoForm(Core.Load(Path.Combine(projectDir, p.Id + ".json"))); Core.Save(draftPath, current); status.Text = "已打开项目"; }); }
        void RefreshProjects() {
            catalog = new List<Project>(); int failures = 0; Directory.CreateDirectory(projectDir);
            foreach (string file in Directory.GetFiles(projectDir, "*.json")) { try { var saved = Core.Load(file); saved.LocalSavedAt = File.GetLastWriteTimeUtc(file); catalog.Add(saved); } catch { failures++; } }
            Filter(); if (failures > 0) MessageBox.Show(failures + " 个项目文件无法读取，已保留原文件。可从同目录 .bak 文件恢复。", "项目读取提示");
        }
        void Filter() { projects.Items.Clear(); foreach (var group in Core.GroupProjects(catalog, current.Id).Where(g => g.ToString().IndexOf(search.Text, StringComparison.CurrentCultureIgnoreCase) >= 0)) { projects.Items.Add(group); if (group.Versions.Any(p => p.Id == current.Id)) projects.SelectedItem = group; } }
        void ImportProject() {
            using (var d = new OpenFileDialog { Filter = "项目文件 (*.json)|*.json", Title = "导入项目副本" }) if (d.ShowDialog() == DialogResult.OK) Run(() => {
                var p = Core.Load(d.FileName); if (!LeaveCurrent()) return; p.Id = Guid.NewGuid().ToString("N");
                if (!string.IsNullOrEmpty(p.Template) && !File.Exists(Core.ResolveTemplate(p.Template))) { MessageBox.Show("关联模板不存在，已切换到内置结构。请重新导入模板。"); p.Template = ""; }
                LoadIntoForm(p); SaveProject();
            });
        }
        byte[] CurrentTemplate() { if (selectedTemplate == "") return Core.BuiltIn(); string path = Core.ResolveTemplate(selectedTemplate); if (!File.Exists(path)) throw new FileNotFoundException("关联模板找不到，请重新导入模板或使用内置结构。"); return File.ReadAllBytes(path); }
        byte[] CurrentWordingTemplate() { if (selectedTemplate == "") return Core.SourceTemplate(); string path = Core.ResolveTemplate(selectedTemplate); string source = path + ".source.docx"; return File.Exists(source) ? File.ReadAllBytes(source) : File.ReadAllBytes(path); }
        UiPreferences LoadUiPreferences() { try { if (!File.Exists(uiPreferencesPath)) return new UiPreferences(); using (var stream = File.OpenRead(uiPreferencesPath)) return (UiPreferences)new DataContractJsonSerializer(typeof(UiPreferences)).ReadObject(stream) ?? new UiPreferences(); } catch { return new UiPreferences(); } }
        string LastExportDirectory() { string path = LoadUiPreferences().LastExportDirectory; return !string.IsNullOrWhiteSpace(path) && Directory.Exists(path) ? path : Path.Combine(Core.Root, "reports"); }
        void RememberExportDirectory(string fileName) { try { var value = new UiPreferences { LastExportDirectory = Path.GetDirectoryName(Path.GetFullPath(fileName)) }; using (var memory = new MemoryStream()) { new DataContractJsonSerializer(typeof(UiPreferences)).WriteObject(memory, value); Core.Atomic(uiPreferencesPath, memory.ToArray()); } } catch { } }
        void ImportTemplate() {
            using (var d = new OpenFileDialog { Filter = "Word 模板 (*.docx)|*.docx", Title = "导入带字段标记的Word模板" }) if (d.ShowDialog() == DialogResult.OK) Run(() => {
                byte[] source; using (var stream = new FileStream(d.FileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) using (var memory = new MemoryStream()) { stream.CopyTo(memory); source = memory.ToArray(); } byte[] bytes = Core.PrepareTemplate(source);
                if (Core.HasTemplateFields(source) && File.Exists(d.FileName + ".source.docx")) source = File.ReadAllBytes(d.FileName + ".source.docx");
                if (!Core.HasTemplateFields(bytes)) {
                    using (var mapping = new TemplateMappingForm(bytes)) { if (mapping.ShowDialog(this) != DialogResult.OK) return; bytes = mapping.Result; }
                }
                var fields = Core.ValidateTemplate(bytes);
                string dest = Path.Combine(Core.Data, "templates", DateTime.Now.ToString("yyyyMMdd_HHmmss") + "_" + Guid.NewGuid().ToString("N").Substring(0, 6) + "_" + Core.SafeName(Path.GetFileNameWithoutExtension(d.FileName)) + ".docx");
                Core.Atomic(dest + ".source.docx", source); Core.Atomic(dest, bytes); ApplyTemplateToForm(source, bytes, dest.Substring(Core.Root.Length)); RefreshPreview(); status.Text = "模板已导入，项目信息、工期及有关事项已刷新；导入前项目已留存副本。";
            });
        }
        void EditTemplateMappings() {
            Run(() => {
                byte[] original = CurrentTemplate();
                if (selectedTemplate == "") original = Core.SourceTemplate();
                else { string sourcePath = Core.ResolveTemplate(selectedTemplate) + ".source.docx"; if (File.Exists(sourcePath)) original = File.ReadAllBytes(sourcePath); }
                using (var mapping = new TemplateMappingForm(Core.PrepareTemplate(CurrentTemplate()))) {
                    if (mapping.ShowDialog(this) != DialogResult.OK) return;
                    string path = Path.Combine(Core.Data, "templates", "填写位置_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + "_" + Guid.NewGuid().ToString("N").Substring(0, 6) + ".docx");
                    Gather(); Core.Atomic(path + ".source.docx", original); Core.Atomic(path, mapping.Result); current.Template = path.Substring(Core.Root.Length); Core.ApplyTemplateDefaults(current, original, mapping.Result); LoadIntoForm(current); Changed(); RefreshPreview();
                }
            });
        }
        void ExportTemplate() {
            using (var d = new SaveFileDialog { Filter = "Word 文档 (*.docx)|*.docx", FileName = "结算审核_字段模板.docx", OverwritePrompt = true }) if (d.ShowDialog() == DialogResult.OK) Run(() => { Core.Atomic(d.FileName, Core.BuiltIn()); status.Text = "字段模板已导出。可将其中的字段标记放进自己的 .docx 模板副本。"; });
        }
        void LayoutWordPreview() { int margin = (int)(20 * DeviceDpi / 96F); int width = Math.Min((int)(850 * DeviceDpi / 96F), Math.Max(420, wordPreviewHost.ClientSize.Width - margin * 2)); preview.SetBounds(Math.Max(margin, (wordPreviewHost.ClientSize.Width - width) / 2), margin, width, Math.Max(200, wordPreviewHost.ClientSize.Height - margin * 2)); }
        void RefreshPreview() { try { Gather(); WordPreviewRenderer.Render(preview, Core.Export(current, CurrentTemplate(), true)); LayoutWordPreview(); } catch (Exception ex) { preview.Clear(); preview.Font = new Font("Microsoft YaHei UI", 11F); preview.Text = "暂时无法预览：\r\n" + ex.Message; } }
        void ExportReport(bool draft) {
            Run(() => {
                excelGrid.EndEdit(); Gather(); if (!draft && current.Excel != null && current.Excel.AutoLink && LocateExcelProblems()) return;
                if (!draft && current.Excel != null && current.Excel.AutoLink && ExcelReview.BalanceIssues(current.Excel).Count > 0) { ShowBalanceDialog(); status.Text = "请处理工程总额与差额中的红色提示后再生成。"; return; }
                byte[] template = CurrentTemplate(); byte[] output = Core.Export(current, template, draft);
                string name = Core.SafeName(Core.Get(current, "项目名称")) + "_结算审核" + (draft ? "草稿" : "审核稿") + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".docx";
                using (var d = new SaveFileDialog { Filter = "可编辑 Word 文档 (*.docx)|*.docx", FileName = name, InitialDirectory = LastExportDirectory(), RestoreDirectory = true, OverwritePrompt = true, Title = "选择审核报告保存位置" }) {
                    Directory.CreateDirectory(d.InitialDirectory); if (d.ShowDialog(this) != DialogResult.OK) return;
                    if (File.Exists(d.FileName)) throw new IOException("为保留历史版本，请换一个尚不存在的文件名。");
                    SaveProject();
                    string history = Path.Combine(Core.Data, "history", DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + "_" + Guid.NewGuid().ToString("N").Substring(0, 6));
                    Directory.CreateDirectory(history); Core.Save(Path.Combine(history, "project.json"), current); File.WriteAllBytes(Path.Combine(history, "template.docx"), template);
                    File.WriteAllBytes(Path.Combine(history, "report.docx"), output);
                    using (var file = new FileStream(d.FileName, FileMode.CreateNew)) file.Write(output, 0, output.Length);
                    RememberExportDirectory(d.FileName);
                    File.WriteAllText(Path.Combine(history, "导出记录.txt"), "版本：" + Core.Version + "\r\n导出时间：" + DateTime.Now.ToString("O") + "\r\n类型：" + (draft ? "草稿" : "审核稿") + "\r\n输出位置：" + d.FileName, Encoding.UTF8);
                    status.Text = "Word 已生成：" + d.FileName;
                    if (MessageBox.Show("已生成可编辑Word。请在Word或WPS中复核分页及版式。\n\n" + d.FileName + "\n\n现在打开？", "导出完成", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes) Core.Open(d.FileName);
                }
            });
        }
        public void Screenshot(string path, int tab) { tabs.SelectedIndex = tab; Show(); Application.DoEvents(); using (var bmp = new Bitmap(Width, Height)) { DrawToBitmap(bmp, new Rectangle(0, 0, Width, Height)); bmp.Save(path); } Hide(); }
    }
    static class Program {
        [STAThread] static int Main(string[] args) {
            Application.EnableVisualStyles(); Application.SetCompatibleTextRenderingDefault(false);
            try {
                if (args.Length > 0 && args[0] == "--test") return Tests.Run(args[1]);
                string mutexId;
                using (var sha = System.Security.Cryptography.SHA256.Create()) mutexId = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(Core.Root.ToUpperInvariant()))).Replace("-", "");
                bool created;
                using (var mutex = new System.Threading.Mutex(true, "Local\\SettlementReview_" + mutexId, out created)) {
                    if (!created) { MessageBox.Show("这个目录的软件已经打开，请回到已打开的窗口继续。", "结算审核工作台"); return 0; }
                    try { Application.Run(new MainForm()); } finally { mutex.ReleaseMutex(); }
                } return 0;
            } catch (Exception ex) { if (args.Length > 0) { File.WriteAllText(Path.Combine(args[1], "failure.txt"), ex.ToString()); return 1; } MessageBox.Show("启动失败：" + ex.Message + "\n请把完整软件目录放在可写位置后重试。", "结算审核工作台"); return 1; }
        }
    }
}
