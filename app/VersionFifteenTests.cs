using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.IO.Compression;
using System.Windows.Forms;
using System.Xml.Linq;
namespace SettlementReview {
    public static partial class Tests {
        static void CheckVersionFifteen(string output) {
            var project = new Project(); foreach (var f in Core.Body) project.Values[f.Name] = "无";
            var values = Core.Values(project, true);
            Assert(values["有关事项说明"] == "" && values["修正造价情况说明"] == "", "无事项汇总不生成空标题和无字");
            byte[] report = Core.Export(project, Core.BuiltIn(), true); string path = Path.Combine(output,"v015_无事项报告.docx"); File.WriteAllBytes(path, report);
            using (var stream = new MemoryStream(report)) using (var zip = new ZipArchive(stream, ZipArchiveMode.Read)) {
                string text = string.Concat(XDocument.Load(zip.GetEntry("word/document.xml").Open()).Descendants(Core.W + "t").Select(t => t.Value));
                Assert(!text.Contains("工期延期报告") && !text.Contains("标化工地说明") && !text.Contains("合同标内调整"), "真实导出无事项正文及标题删除");
                Assert(!text.Contains("合同标外工程量及综合单价") && !text.Contains("联系单工程量及综合单价") && !text.Contains("合同标内工程量及综合单价"), "原始长标题在事项为无时同步删除");
            }
            project.Values["合同标外调整说明"] = "保留的标外审核内容";
            byte[] retained = Core.Export(project, Core.BuiltIn(), true);
            string retainedText = Read(retained, "word/document.xml");
            Assert(retainedText.Contains("合同标外工程量及综合单价") && retainedText.Contains("保留的标外审核内容"), "有实际内容的长标题保留");
            project.Values["合同标外调整说明"] = "无";
            project.Values["工期延期报告"] = "延期批准十天";
            report = Core.Export(project, Core.BuiltIn(), true);
            using (var stream = new MemoryStream(report)) using (var zip = new ZipArchive(stream, ZipArchiveMode.Read)) {
                string text = string.Concat(XDocument.Load(zip.GetEntry("word/document.xml").Open()).Descendants(Core.W + "t").Select(t => t.Value));
                Assert(text.Contains("延期批准十天"), "有内容的事项保持输出");
            }
            string isolated = Path.Combine(output,"v015_data"); Directory.CreateDirectory(isolated);
            using (var form = new MainForm(true, isolated)) {
                var library = new List<BasisTemplate>(); BasisTemplateStore.Save(Path.Combine(isolated,"basis-templates.json"), library,"a","新依据","新增依据条目");
                string choices = (string)V7Call(form,"BasisChoices","当前手动条目","旧模板条目");
                Assert(choices.Contains("新增依据条目") && choices.Contains("当前手动条目"), "选择依据即时读取个人库及当前编辑内容");
                BasisTemplateStore.Save(Path.Combine(isolated,"basis-templates.json"), library,"a","改名依据","修改依据条目");
                choices = (string)V7Call(form,"BasisChoices","当前手动条目","旧模板条目");
                Assert(choices.Contains("修改依据条目") && !choices.Contains("新增依据条目"), "依据库修改后选择列表实时更新");
                V7Call(form,"LoadIntoForm",project);
                string file = Path.Combine(output,"v015_自选位置项目.json"); V7Call(form,"SaveProjectFile",file);
                Assert(File.Exists(file) && Core.Get(Core.Load(file),"工期延期报告") == "延期批准十天", "自选位置保存保留正文");
                V7Call(form,"OpenProjectPath",file);
                Assert(Core.Get(V7Field<Project>(form,"current"),"工期延期报告") == "延期批准十天", "按路径打开项目保留编辑数据");
                string originalTemplate = Path.Combine(isolated,"imported.docx"); File.WriteAllBytes(originalTemplate, Core.BuiltIn()); project.Template = originalTemplate;
                V7Call(form,"LoadIntoForm",project); V7Call(form,"SaveProjectFile",file);
                Assert(Core.Load(file).Template == "v015_自选位置项目.template.docx", "项目文件保存相对模板关联便于迁移");
                string moved = Path.Combine(output,"v015_另一电脑目录"); Directory.CreateDirectory(moved);
                foreach (var source in Directory.GetFiles(output,"v015_自选位置项目*")) File.Copy(source,Path.Combine(moved,Path.GetFileName(source)),true);
                V7Call(form,"OpenProjectPath",Path.Combine(moved,Path.GetFileName(file)));
                Assert(V7Field<Project>(form,"current").Template.StartsWith(moved) && ((byte[])V7Call(form,"CurrentTemplate")).Length > 0, "迁移另一目录打开时恢复相邻报告模板");
                form.Show(); Application.DoEvents(); form.Screenshot(Path.Combine(output,"v015_紧凑正文.png"),3); form.Screenshot(Path.Combine(output,"v015_项目保存打开.png"),1);
                var controls = ControlsBelow(form.Controls).ToList();
                Assert(controls.OfType<Button>().Any(b => b.Text == "保存项目到…") && !controls.OfType<Button>().Any(b => b.Text == "打开选中项目"), "左上角文件操作使用明确名称");
                Assert(controls.OfType<System.Windows.Forms.Panel>().Any(t => t.Dock == DockStyle.Top && Math.Abs(t.Height - 48 * form.DeviceDpi / 96F) < 3), "顶部标题栏压缩为48像素"); form.Close();
            }
        }
    }
}
