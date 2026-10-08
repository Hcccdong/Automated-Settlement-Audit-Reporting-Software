using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;

namespace SettlementReview {
    public static partial class Tests {
        static ExcelSheet ReviewTestSheet(string sheetName, string projectName, string itemName, string feature, string sent, string audit) {
            var sheet = new ExcelSheet { Name = sheetName };
            ExcelRow(sheet, 1, "工程审核简明表"); ExcelRow(sheet, 2, "工程名称：", projectName, null, null, null, null, null, null, null, null, null, null, null, null, null, "第1页 共2页");
            ExcelRow(sheet, 3, "项目名称", "项目特征", "送审工程量", "审核工程量", "送审单价", "审核单价", "送审金额", "审核金额");
            ExcelRow(sheet, 4, itemName, feature, "10", "8", "100", "100", sent, audit);
            return sheet;
        }
        static ExcelReviewState CheckVersionSix(string output, Project p, byte[] template) {
            byte[] source = Core.SourceTemplate(); var old = new Project(); old.Values["建设单位"] = "旧建设单位"; old.Values["项目名称"] = "旧项目名称"; old.Values["送审金额"] = "123";
            Core.ApplyTemplateDefaults(old, source, template, true); var defaults = Core.TemplateDefaults(source, template);
            Assert(Core.Get(old, "建设单位") == defaults["建设单位"] && Core.Get(old, "项目名称") == defaults["项目名称"], "重新导入刷新已填写的单位和项目名称");
            Assert(Core.Get(old, "送审金额") == "123" && Core.Get(old, "报告编号") == defaults["报告编号"], "模板刷新显示编号但保留人工审核金额");
            var parts = ReviewTestSheet("一号楼_分部分项", "1号楼", "砌块墙", "外墙200厚\n1、A5.0/B07砌块\n2、Mb5.0砂浆", "1000", "800");
            var measures = ReviewTestSheet("任意表名_技术措施", "1号楼", "模板", "木模板\n1、层高3.6m", "1000", "1200");
            var moreParts = ReviewTestSheet("一号楼分部分项续表", "1号楼", "混凝土", "C30\n1、泵送", "1000", "900");
            var second = ReviewTestSheet("另一个工作表_分部分项", "2号楼", "砌块墙", "内墙150厚\n1、A5.0/B07砌块", "1000", "750");
            var state = ExcelReview.Parse(new List<ExcelSheet> { parts, measures, moreParts, second }, "分部分项及技术措施示例.xlsx");
            Assert(state.Items[0].Group == "1号楼" && state.Items[1].Group == "1号楼" && state.Items[3].Group == "2号楼", "拆分单元格工程名称优先于工作表标签且去除页码");
            Assert(state.Items[0].Category == "分部分项" && state.Items[1].Category == "技术措施", "识别两类费用而非两个子单位工程");
            string text = ExcelReview.Compose(state, true);
            Assert(Regex.Matches(text, "1号楼：").Count == 1 && !text.Contains("分部分项：") && !text.Contains("技术措施：") && text.IndexOf("混凝土") < text.IndexOf("2号楼："), "同名子单位合并且取消费用分类标题");
            Assert(text.Contains("1. 砌块墙") && text.Contains("3. 混凝土") && text.Contains("2. 模板") && text.Contains("4. 砌块墙"), "同一调整章节从1开始跨分类连续编号");
            Assert(!text.Contains("编码") && !text.Contains("项目特征") && !text.Contains("（核增）") && !text.Contains("（核减）") && !text.Contains("任意表名"), "生成文字去除编码标签工作表名及核增减括号");
            Assert(text.Contains("砌块墙，送审金额为1000.00元，审核金额为800.00元") && !text.Contains("A5.0/B07砌块") && text.Contains("此项正偏差200.00元"), "名称无歧义时省略项目特征并使用指定金额句式");
            Assert(ExcelReview.Describe("防水", "1.5厚涂膜\n1、上翻300mm").Contains("1.5厚涂膜，上翻300mm"), "去特征内部序号时不损坏小数规格");
            var combined = ReviewTestSheet("分部分项", "1#、2#楼", "砌体", "M7.5", "100", "90");
            Assert(ExcelReview.Parse(new List<ExcelSheet> { combined }, "联合楼栋.xlsx").Items.Single().Group == "1#、2#楼", "联合表头名称不擅自拆分金额");
            var mixed = ReviewTestSheet("分部分项与技术措施", "3号楼", "一、分部分项工程", "", "", "");
            ExcelRow(mixed, 5, "墙体", "200厚", "10", "8", "100", "100", "1000", "800");
            ExcelRow(mixed, 6, "二、技术措施项目费");
            ExcelRow(mixed, 7, "脚手架", "搭设高度6m", "10", "9", "100", "100", "1000", "900");
            var mixedState = ExcelReview.Parse(new List<ExcelSheet> { mixed }, "同表两类.xlsx");
            Assert(mixedState.Items.Count == 2 && mixedState.Items[0].Category == "分部分项" && mixedState.Items[1].Category == "技术措施", "同一工作表内的两类分节可分别识别且不混入分节汇总行");
            p.Values["合同标内调整说明"] = "先前人工正文"; ExcelReview.StartLink(p, state);
            Assert(Core.Get(p, "合同标内调整说明") == text && state.OriginalBodies["合同标内调整说明"] == "先前人工正文", "导入即联动并保存联动前正文");
            state.Items[0].Reason = "单价偏差"; ExcelReview.Synchronize(p, false); Assert(Core.Get(p, "合同标内调整说明").Contains("由于单价偏差"), "修改原因后正文实时联动");
            state.Items[0].Reason = "工程量偏差"; state.Items[3].Target = "合同标外调整说明"; ExcelReview.Synchronize(p, false);
            Assert(!Core.Get(p, "合同标内调整说明").Contains("2号楼") && Core.Get(p, "合同标外调整说明").Contains("1. 砌块墙"), "调整报告归属后两章节同步且分别从1编号");
            state.Items[0].Include = false; ExcelReview.Synchronize(p, false); Assert(!Core.Get(p, "合同标内调整说明").Contains("外墙200厚") && Core.Get(p, "合同标内调整说明").Contains("1. 模板"), "取消采用后移除旧明细并重新编号"); state.Items[0].Include = true;
            state.Items[1].Reason = "待确认"; Assert(Core.Preview(p, template).Contains("调整原因待确认"), "待确认原因仍可在联动草稿预览"); Reject(() => Core.Export(p, template, false), "待确认项目不混入正式生成稿"); state.Items[1].Reason = "工程量偏差";
            byte[] word = Core.Export(p, template, false); var document = XDocument.Parse(Read(word, "word/document.xml")); var numberDoc = XDocument.Parse(Read(word, "word/numbering.xml"));
            HashSet<string> ids; using (var ms = new MemoryStream(word)) using (var zip = new ZipArchive(ms, ZipArchiveMode.Read)) ids = Core.ReviewNumberIds(zip);
            var generated = document.Descendants(Core.W + "p").Where(par => par.Element(Core.W + "pPr") != null && par.Element(Core.W + "pPr").Element(Core.W + "numPr") != null && ids.Contains((string)par.Element(Core.W + "pPr").Element(Core.W + "numPr").Element(Core.W + "numId").Attribute(Core.W + "val"))).ToList();
            Assert(generated.Count == 4 && ids.Count == 2, "明细变为四个独立Word列表段落并分章节建序列");
            Assert(generated.All(par => !par.Descendants(Core.W + "br").Any() && !Regex.IsMatch(Core.RawParagraphText(par), @"^\d+\.")), "使用原生Word编号而非重复文字序号或软换行");
            Assert(numberDoc.Root.Elements(Core.W + "num").Where(n => ids.Contains((string)n.Attribute(Core.W + "numId"))).All(n => (string)n.Element(Core.W + "lvlOverride").Element(Core.W + "startOverride").Attribute(Core.W + "val") == "1"), "每个原生编号序列明确从1开始");
            var templateDoc = XDocument.Parse(Read(template, "word/document.xml")); var donor = templateDoc.Descendants(Core.W + "p").First(par => Core.RawParagraphText(par).Contains("{{基期价格}}"));
            Assert(generated.All(par => XNode.DeepEquals(par.Element(Core.W + "pPr").Element(Core.W + "spacing"), donor.Element(Core.W + "pPr").Element(Core.W + "spacing"))), "每条Excel说明沿用有关事项段落行距");
            var originalNumber = XDocument.Parse(Read(template, "word/numbering.xml")); Assert(originalNumber.Root.Elements().All(n => numberDoc.Root.Elements().Any(other => XNode.DeepEquals(n, other))), "原模板编号定义完整保留不干扰其他章节");
            Assert(Core.Preview(p, template).Contains("1. 砌块墙") && Core.Preview(p, template).Contains("2. 模板"), "报告导出页文字预览显示原生编号");
            string compactXml = "<w:document xmlns:w='" + Core.W + "'><w:body><w:p><w:r><w:t>{{项目名称}}</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:line='480' w:lineRule='exact'/></w:pPr><w:r><w:t>{{修正造价情况说明}}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>";
            byte[] compact = Change(template, "word/document.xml", compactXml);
            var compactDoc = XDocument.Parse(Read(Core.Export(p, compact, false), "word/document.xml"));
            Assert(compactDoc.Descendants(Core.W + "numPr").Count() == 4 && !compactDoc.Descendants(Core.W + "br").Any(), "自定义模板汇总填写位置也生成独立编号段落");
            byte[] missingTarget = Change(template, "word/document.xml", compactXml.Replace("{{修正造价情况说明}}", "固定说明"));
            Reject(() => Core.Export(p, missingTarget, false), "模板缺少联动正文位置时防止静默遗漏明细");
            byte[] altered = Change(template, "word/document.xml", templateDoc.ToString().Replace("w:line=\"520\"", "w:line=\"600\""));
            var alteredDoc = XDocument.Parse(Read(Core.Export(p, altered, false), "word/document.xml")); Assert(alteredDoc.Descendants(Core.W + "p").Where(par => Core.RawParagraphText(par).Contains("此项负偏差200.00元")).All(par => (string)par.Element(Core.W + "pPr").Element(Core.W + "spacing").Attribute(Core.W + "line") == "600"), "换模板后行距跟随当前模板而非硬编码");
            File.WriteAllBytes(Path.Combine(output, "v060_联动分组编号验证.docx"), word); File.WriteAllText(Path.Combine(output, "v060_组稿文字.txt"), Core.Preview(p, template), Encoding.UTF8);
            string save = Path.Combine(output, "v060_联动保存.json"); Core.Save(save, p); var reloaded = Core.Load(save); reloaded.Excel.Items[0].Note = "复核图纸"; ExcelReview.Synchronize(reloaded, false); Assert(Core.Get(reloaded, "合同标内调整说明").Contains("复核图纸") && reloaded.Excel.Items[1].Category == "技术措施", "重新打开项目仍保留分类归属及联动关系");
            using (var form = new MainForm(true)) {
                var load = typeof(MainForm).GetMethod("LoadIntoForm", BindingFlags.NonPublic | BindingFlags.Instance); load.Invoke(form, new object[] { old });
                typeof(MainForm).GetMethod("ApplyTemplateToForm", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(form, new object[] { source, template, "" });
                var inputs = (Dictionary<string, TextBox>)typeof(MainForm).GetField("inputs", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(form);
                Assert(inputs["项目名称"].Text == defaults["项目名称"] && inputs["送审金额"].Text == "123", "实际导入界面刷新项目输入框而非只更新后台值");
                load.Invoke(form, new object[] { p });
                Assert(inputs["合同标内调整说明"].ReadOnly && inputs["合同标内调整说明"].Text.Contains("1号楼："), "正文界面显示Excel联动结果并避免两处同时编辑");
                state.Items[0].Note = "界面联动验证"; typeof(MainForm).GetMethod("Changed", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(form, null);
                Assert(inputs["合同标内调整说明"].Text.Contains("界面联动验证"), "明细编辑事件将最新文字同步回正文窗口");
            }
            state.Items[0].Note = ""; ExcelReview.Synchronize(p, false);
            return state;
        }
    }
}
