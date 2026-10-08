using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;

namespace SettlementReview {
    public static partial class Core {
        public static byte[] RemoveAbsentSections(byte[] bytes, Project project, Dictionary<string,string> values) {
            var absent = new HashSet<string>(Body.Where(f => f.Name != "审核依据及原则" && IsAbsent(values[f.Name])).Select(f => f.Name));
            if (string.IsNullOrWhiteSpace(values["有关事项说明"])) absent.Add("有关事项说明");
            if (string.IsNullOrWhiteSpace(values["修正造价情况说明"])) absent.Add("修正造价情况说明");
            using (var input = new MemoryStream(bytes)) using (var source = new ZipArchive(input, ZipArchiveMode.Read)) using (var output = new MemoryStream()) {
                using (var target = new ZipArchive(output, ZipArchiveMode.Create, true)) foreach (var entry in source.Entries) using (var dest = target.CreateEntry(entry.FullName).Open()) {
                    if (!IsTextPart(entry.FullName)) { using (var stream = entry.Open()) stream.CopyTo(dest); continue; }
                    var xml = Xml(entry); string before = xml.ToString(SaveOptions.DisableFormatting);
                    foreach (var p in xml.Descendants(W + "p").ToList()) {
                        if (p.Parent == null) continue;
                        string raw = RawParagraphText(p);
                        var keys = Regex.Matches(raw, @"\{\{([^{}]+)\}\}").Cast<Match>().Select(m => m.Groups[1].Value).ToList();
                        if (keys.Count == 0 || keys.Any(k => !absent.Contains(k))) continue;
                        var previous = p.ElementsBeforeSelf(W + "p").LastOrDefault();
                        if (previous != null) {
                            string title = Regex.Replace(RawParagraphText(previous).Trim(), @"^[一二三四五六七八九十\d\s、.．（）()]+", "").TrimEnd('：', ':');
                            if (keys.Any(k => AbsentSectionHeading(title, k))) PreserveSectionOrRemove(previous);
                        }
                        PreserveSectionOrRemove(p);
                    }
                    if (before == xml.ToString(SaveOptions.DisableFormatting)) { using (var stream = entry.Open()) stream.CopyTo(dest); } else using (var writer = new StreamWriter(dest, new UTF8Encoding(false))) xml.Save(writer, SaveOptions.DisableFormatting);
                }
                return output.ToArray();
            }
        }
        static bool AbsentSectionHeading(string title, string key) {
            if (title == key || title == key.Replace("说明", "") || key == "修正造价情况说明" && title == "修正情况分析") return true;
            string prefix = key == "合同标内调整说明" ? "合同标内" : key == "合同标外调整说明" ? "合同标外" : key == "联系单调整说明" ? "联系单" : null;
            return prefix != null && Regex.IsMatch(title, "^" + prefix + @"(?:工程量及综合单价|工程量及综合单价重新计算|工程量及综合单价重新计算调整|工程量及综合单价调整|调整)(?:部分)?$");
        }
        static void PreserveSectionOrRemove(XElement p) {
            var section = p.Descendants(W + "sectPr").FirstOrDefault();
            if (section != null) { p.ReplaceNodes(new XElement(W + "pPr", new XElement(section))); return; }
            if (p.Parent.Name == W + "tc" && p.Parent.Elements(W + "p").Count() == 1) p.RemoveNodes(); else p.Remove();
        }
    }
    public partial class MainForm {
        string projectFilePath = "";
        string BasisChoices(string active, string original) {
            string personal = Core.Get(current, "当前依据选择内容");
            var texts = new List<string> { active, personal };
            texts.AddRange(BasisTemplateStore.Load(BasisLibraryPath).Select(t => t.Text)); texts.Add(original);
            return string.Join("\r\n", texts.SelectMany(t => (t ?? "").Replace("\r\n", "\n").Split('\n')).Where(t => !string.IsNullOrWhiteSpace(t)).Distinct());
        }
        void SaveProjectAs() { Run(() => {
            Gather();
            using (var dialog = new SaveFileDialog { Title = "选择项目保存位置", Filter = "结算审核项目 (*.json)|*.json", DefaultExt = "json", AddExtension = true, FileName = projectFilePath == "" ? Core.SafeName(Core.Get(current, "项目名称")) + ".json" : Path.GetFileName(projectFilePath), InitialDirectory = projectFilePath == "" ? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) : Path.GetDirectoryName(projectFilePath) }) {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                SaveProjectFile(dialog.FileName);
            }
        }); }
        void SaveProjectFile(string path) {
            Gather(); string original = current.Template;
            try {
                if (!string.IsNullOrEmpty(selectedTemplate)) {
                    string templateFile = Path.GetFileNameWithoutExtension(path) + ".template.docx";
                    string dest = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(path)), templateFile);
                    Core.Atomic(dest, CurrentTemplate()); Core.Atomic(dest + ".source.docx", CurrentWordingTemplate()); current.Template = templateFile;
                }
                Core.Save(path, current);
            } finally { current.Template = original; }
            projectFilePath = Path.GetFullPath(path); SaveProject(); status.Text = "项目已保存：" + projectFilePath + "（换电脑时请一起复制同名模板文件）";
        }
        void OpenProjectFile() { Run(() => {
            using (var dialog = new OpenFileDialog { Title = "打开结算审核项目", Filter = "结算审核项目 (*.json)|*.json" }) {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                OpenProjectPath(dialog.FileName);
            }
        }); }
        void OpenProjectPath(string path) {
            var p = Core.Load(path);
            if (!string.IsNullOrEmpty(p.Template) && !Path.IsPathRooted(p.Template)) {
                string adjacent = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(Path.GetFullPath(path)), p.Template));
                if (File.Exists(adjacent)) p.Template = adjacent;
            }
            if (!string.IsNullOrEmpty(p.Template) && !File.Exists(Core.ResolveTemplate(p.Template))) throw new FileNotFoundException("找不到此项目的报告模板，请将同名 .template.docx 及 .source.docx 一起复制过来。");
            if (!LeaveCurrent()) return;
            LoadIntoForm(p); projectFilePath = Path.GetFullPath(path); Core.Save(draftPath, current); status.Text = "已打开：" + projectFilePath;
        }
    }
}
