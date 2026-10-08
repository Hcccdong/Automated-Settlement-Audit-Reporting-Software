using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;

namespace SettlementReview {
    public static partial class Tests {
        static IEnumerable<Control> ControlsBelow(Control.ControlCollection controls) { foreach (Control c in controls) { yield return c; foreach (Control child in ControlsBelow(c.Controls)) yield return child; } }
        static void ExcelRow(ExcelSheet s, int row, params string[] cells) { s.Rows[row] = cells.Select((value, i) => new { value, i }).Where(x => x.value != null).ToDictionary(x => x.i + 1, x => x.value); }
        static string ColumnLetter(int number) { string s = ""; while (number > 0) { number--; s = (char)('A' + number % 26) + s; number /= 26; } return s; }
        static void WriteExcelFixture(ExcelSheet sheet, string path) {
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var rows = new XElement(ns + "sheetData", sheet.Rows.Select(row => new XElement(ns + "row", new XAttribute("r", row.Key), row.Value.Select(cell => new XElement(ns + "c", new XAttribute("r", ColumnLetter(cell.Key) + row.Key), new XAttribute("t", "inlineStr"), new XElement(ns + "is", new XElement(ns + "t", cell.Value)))))));
            var merges = new XElement(ns + "mergeCells", sheet.Merges.Select(m => new XElement(ns + "mergeCell", new XAttribute("ref", ColumnLetter(m[1]) + m[0] + ":" + ColumnLetter(m[3]) + m[2]))));
            using (var file = File.Create(path)) using (var zip = new ZipArchive(file, ZipArchiveMode.Create)) {
                var parts = new Dictionary<string, string> {
                    { "xl/workbook.xml", "<workbook xmlns='" + ns + "' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><sheets><sheet name='测试分部工程' sheetId='1' r:id='rId1'/></sheets></workbook>" },
                    { "xl/_rels/workbook.xml.rels", "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Target='worksheets/sheet1.xml'/></Relationships>" },
                    { "xl/worksheets/sheet1.xml", new XElement(ns + "worksheet", rows, merges).ToString() }
                };
                foreach (var pair in parts) using (var w = new StreamWriter(zip.CreateEntry(pair.Key).Open(), new UTF8Encoding(false))) w.Write(pair.Value);
            }
        }
        static ExcelReviewState CheckExcel(string output, Project reportProject, byte[] template) {
            var sheet = new ExcelSheet { Name = "测试分部工程" };
            ExcelRow(sheet, 1, "工程审核简明表"); ExcelRow(sheet, 2, "工程名称：审核测试楼");
            ExcelRow(sheet, 3, "序号", "送审", null, null, null, null, null, null, "审查", null, null, null, null, null, null, "核增", "核减");
            ExcelRow(sheet, 4, null, "编号", "名    称", "项目特征", "单位", "数量", "单价", "合价", "编号", "名    称", "项目特征", "单位", "数量", "单价", "合价");
            sheet.Merges.Add(new[] { 3, 2, 3, 8 }); sheet.Merges.Add(new[] { 3, 9, 3, 15 }); sheet.Merges.Add(new[] { 3, 16, 4, 16 }); sheet.Merges.Add(new[] { 3, 17, 4, 17 }); sheet.Merges.Add(new[] { 3, 1, 4, 1 });
            string feature = "外墙200厚\n1、蒸压加气砼砌块A5.0/B07\n2、Mb5.0专用砂浆实砌，完整规格不得截断。";
            ExcelRow(sheet, 5, "1", "010402001002", "砌块墙", feature, "m3", "10", "100", "1000", "010402001002", "砌块墙", feature, "m3", "8", "100", "800", null, "200");
            ExcelRow(sheet, 6, "2", "010402001003", "砌块墙", "外墙150厚；砂浆强度M7.5", "m3", "10", "100", "1000", "010402001003", "砌块墙", "外墙150厚；砂浆强度M7.5", "m3", "10", "120", "1200", "200", null);
            ExcelRow(sheet, 7, "3", "A3", "混凝土", "C30", "m3", "10", "100", "1000", "A3", "混凝土", "C30", "m3", "8", "120", "960", null, "40");
            ExcelRow(sheet, 8, "4", "A4", "钢筋", "HRB400", "t", "1", "100", "100", "A4", "钢筋", "HRB500", "t", "1", "120", "120", "20", null);
            ExcelRow(sheet, 9, "5", "A5", "模板", "木模", "m2", null, "20", "100", "A5", "模板", "木模", "m2", "0", "20", "0", null, "100");
            ExcelRow(sheet, 10, "6", "A6", "不变项目", "相同规格", "m2", "1", "10", "10", "A6", "不变项目", "相同规格", "m2", "1", "10", "10");
            ExcelRow(sheet, 11, null, "0104", "分部汇总", null, null, null, null, "3210", null, null, null, null, null, null, "3090", null, "120");
            string path = Path.Combine(output, "Excel读取测试夹具.xlsx"); WriteExcelFixture(sheet, path);
            var state = ExcelReview.Parse(ExcelReview.Read(path), "测试对比表.xlsx");
            Assert(state.Items.Count == 6, "读取合并表头并排除无序号分部汇总");
            Assert(state.Items[0].Delta == -200 && state.Items[0].Reason == "工程量偏差", "工程量核减负偏差");
            Assert(state.Items[1].Delta == 200 && state.Items[1].Reason == "单价偏差", "单价核增正偏差");
            Assert(state.Items[0].DecreaseAmount == 200 && state.Items[0].IncreaseAmount == 0 && state.Items[1].IncreaseAmount == 200 && state.Items[1].DecreaseAmount == 0, "核增核减分列方向一致");
            Assert(state.Items[2].Reason == "工程量及单价偏差", "量价同时变化");
            Assert(!state.Items[3].Include && state.Items[3].Warning.Contains("项目特征不一致"), "特征不一致默认不采用");
            Assert(state.Items[4].Reason == "待确认" && state.Items[4].SentQuantity == null && state.Items[4].AuditQuantity == 0, "缺失数量不当作零或推断量价原因");
            Assert(state.Items[0].Feature.Contains("完整规格不得截断") && state.Items[1].Name == state.Items[0].Name && state.Items[0].Code != state.Items[1].Code, "同名不同规格不合并且完整保留特征");
            Assert(!state.Items[5].Include, "无偏差项目不默认组稿");
            state.Items[4].Include = false;
            string text = ExcelReview.Compose(state, true);
            Assert(text.Contains("负偏差200.00元") && text.Contains("正偏差200.00元") && text.Contains("200厚") && !text.Contains("完整规格不得截断"), "同名项目只保留最少区别特征");
            Assert(!text.Contains("取费基数变化") && ExcelReview.Summary(state).Contains("净核减 40.00 元"), "不凭空配平或推断费用原因");
            state.Items[3].Include = true; Reject(() => ExcelReview.Compose(state, true), "特征冲突必须核实后写入");
            state.Items[3].Confirmed = true; Assert(ExcelReview.Compose(state, true).Contains("审核为钢筋，HRB500"), "确认差异后保留两侧特征而非强行同名"); state.Items[3].Include = false;
            state.Threshold = 201; Reject(() => ExcelReview.Compose(state, true), "阈值排除全部时不写空段落"); state.Threshold = 200; Assert(ExcelReview.Selected(state).Count == 2, "阈值等值列入"); state.Threshold = 0;
            state.Detailed = true; Assert(ExcelReview.Compose(state, true).Contains("送审工程量10m3、单价100元"), "量价明细句式可选"); state.Detailed = false;
            state.Items[0].Reason = ExcelReview.Reasons[3]; Reject(() => ExcelReview.Compose(state, true), "其他原因需要说明"); state.Items[0].Note = "现场核实墙体工程量"; Assert(ExcelReview.Compose(state, true).Contains("现场核实墙体工程量"), "人工原因可编辑"); state.Items[0].Reason = ExcelReview.Reasons[0];
            var generic = new ExcelSheet { Name = "通用表" };
            ExcelRow(generic, 1, "项目名称", "项目特征", "送审工程量", "审核工程量", "送审单价", "审核单价", "送审金额", "审核金额", "核增", "核减");
            ExcelRow(generic, 2, "路面", "C30混凝土", "10", "9", "100", "100", "1000", "900", "0", "90");
            ExcelRow(generic, 3, "合计", "", "", "", "", "", "1000", "900", "0", "90");
            var genericState = ExcelReview.Parse(new List<ExcelSheet> { generic }, "通用.xlsx");
            Assert(genericState.Items.Count == 1 && genericState.Items[0].Warning.Contains("不一致") && genericState.Items[0].Delta == -100, "通用表核增减不平衡保留来源并提示");
            generic.Rows[2][7] = "#公式未缓存"; var missingFormula = ExcelReview.Parse(new List<ExcelSheet> { generic }, "公式.xlsx"); Assert(missingFormula.Items[0].Warning.Contains("公式未缓存") && missingFormula.Items[0].SentAmount == null, "公式无缓存不当作零");
            generic.Rows[1][7] = "送审金额（万元）"; Assert(ExcelReview.Parse(new List<ExcelSheet> { generic }, "万元.xlsx").Items.Count == 0, "不把万元误当作元");
            Reject(() => ExcelReview.Read("旧表.xls"), "旧式xls给出另存提示");
            var project = new Project { Excel = state }; project.Values["合同标内调整说明"] = "人工原有说明。";
            text = ExcelReview.Compose(state, true); ExcelReview.Apply(project, "合同标内调整说明", text); ExcelReview.Apply(project, "合同标内调整说明", text);
            Assert(project.Values["合同标内调整说明"].IndexOf(text) == project.Values["合同标内调整说明"].LastIndexOf(text) && project.Values["合同标内调整说明"].StartsWith("人工原有说明。"), "写入保留人工原文且重复写入不重复追加");
            state.Items[0].Note = "核对施工图后调整"; string revised = ExcelReview.Compose(state, true); ExcelReview.Apply(project, "合同标内调整说明", revised); Assert(!project.Values["合同标内调整说明"].Contains("现场核实墙体工程量") && project.Values["合同标内调整说明"].Contains("核对施工图后调整"), "重新组稿更新此前生成段");
            string savePath = Path.Combine(output, "Excel项目保存.json"); Core.Save(savePath, project); var loaded = Core.Load(savePath);
            Assert(loaded.Excel.Items.Count == 6 && loaded.Excel.Items[0].Note == "核对施工图后调整" && loaded.Excel.LastText == revised, "明细选择原因及生成记录随项目持久化");
            var bulk = new Project { Excel = new ExcelReviewState { Items = Enumerable.Repeat(state.Items[0], 2500).ToList() } };
            string bulkPath = Path.Combine(output, "批量明细保存.json"); Core.Save(bulkPath, bulk); Assert(Core.Load(bulkPath).Excel.Items.Count == 2500, "2500条明细保存重开不触发默认序列化上限");
            var oversized = new Project(); oversized.Values["超长测试"] = new string('x', 17 * 1024 * 1024); Reject(() => Core.Save(bulkPath, oversized), "过大项目不写成无法重开的文件");
            Assert(Core.Load(bulkPath).Excel.Items.Count == 2500, "超过容量时保留已有项目内容");
            reportProject.Excel = state; reportProject.Values["合同标内调整说明"] = "无"; ExcelReview.Apply(reportProject, "合同标内调整说明", revised);
            byte[] word = Core.Export(reportProject, template, false); Assert(Read(word, "word/document.xml").Contains("200厚") && !Read(word, "word/document.xml").Contains("完整规格不得截断") && Read(word, "word/document.xml").Contains("正偏差200.00元"), "精简后的Excel说明通过当前Word模板导出"); File.WriteAllBytes(Path.Combine(output, "Excel说明导出验证.docx"), word);
            string reference = Path.GetFullPath(Path.Combine(Core.Root, "..", "结算对比表模板.xlsx"));
            if (File.Exists(reference)) { var real = ExcelReview.Parse(ExcelReview.Read(reference), Path.GetFileName(reference)); Assert(real.Items.Count > 0 && real.Items.Any(i => i.Name == "砌块墙" && i.Feature.Contains("A5.0/B07")), "真实用户对比表读取名称与完整特征"); File.WriteAllText(Path.Combine(output, "真实表读取结果.txt"), string.Join("\r\n", real.Messages) + "\r\n" + ExcelReview.Summary(real), Encoding.UTF8); }
            return state;
        }
    }
}
