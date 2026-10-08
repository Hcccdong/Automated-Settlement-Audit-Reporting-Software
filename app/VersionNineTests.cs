using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;

namespace SettlementReview {
    public static partial class Tests {
        static ExcelSheet V9Total(string sheetName, string group, string sent, string audit, string inc = null, string dec = null) {
            var sheet = new ExcelSheet { Name = sheetName };
            ExcelRow(sheet, 1, "专业工程结算审核计算表"); ExcelRow(sheet, 2, "工程名称：" + group);
            ExcelRow(sheet, 3, "序号", "汇总内容", "计算公式", "送审", "审核", "核增", "核减");
            ExcelRow(sheet, 4, "一", "分部分项工程费", "", "999", "899", "", "100");
            ExcelRow(sheet, 12, "八", "建设工程造价", "", sent, audit, inc, dec); return sheet;
        }
        static ExcelReviewState V9Example() {
            var details = ReviewTestSheet("0_1_分部分项", "一号楼", "砌块墙", "200厚", "15000", "6000");
            ExcelRow(details, 5, "防水", "1.5厚涂膜", "10", "8", "100", "100", "1000", "900");
            return ExcelReview.Parse(new List<ExcelSheet> { V9Total("0_1_计算表", "一号楼", "20000", "10000", "", "10000"), details }, "净核减示例.xlsx");
        }
        static decimal V9TextDelta(string text) {
            return Regex.Matches(text, @"此项(正|负)偏差(?:为)?(\d+(?:\.\d+)?)元").Cast<Match>().Sum(m => decimal.Parse(m.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture) * (m.Groups[1].Value == "正" ? 1 : -1));
        }
        static void V9Screenshot(Form dialog, string path) {
            dialog.Show(); Application.DoEvents(); using (var bitmap = new Bitmap(dialog.Width, dialog.Height)) { dialog.DrawToBitmap(bitmap, new Rectangle(0, 0, dialog.Width, dialog.Height)); bitmap.Save(path); } dialog.Hide();
        }
        static void CheckVersionNine(string output, Project example, byte[] template) {
            var state = V9Example(); var balance = state.Balances.Single(); state.Threshold = 500;
            Assert(state.Totals.Count == 1 && state.Items.Count == 2 && balance.Enabled && state.Totals[0].DetailSheets.Single() == "0_1_分部分项", "计算表只读取最终造价行且按编号关联明细");
            var result = ExcelReview.CalculateBalance(state, balance);
            Assert(result.TotalDelta == -10000 && result.Listed == -9000 && result.Omitted == -100 && result.NonDetail == -900 && result.Residual == -1000 && result.Issue == "", "一万元净核减减九千元已列明细得到一千元差额且来源可拆解");
            string text = ExcelReview.Compose(state, true);
            Assert(text.EndsWith("2. " + ExcelReview.FeeReason + "，此项负偏差为1000.00元。") && V9TextDelta(text) == -10000, "按用户原句追加负偏差一千元且报告金额配平");
            state.Threshold = 0; Assert(ExcelReview.Compose(state, true).Contains("3. " + ExcelReview.FeeReason + "，此项负偏差为900.00元。"), "降低门槛后差额随新增明细自动减少");
            state.Items[0].Include = false; Assert(V9TextDelta(ExcelReview.Compose(state, true)) == -10000 && ExcelReview.CalculateBalance(state, balance).Residual == -9900, "取消采用后末条补入对应差额且总额不变");
            state.Threshold = 999999; Assert(ExcelReview.Compose(state, true) == "一号楼：\r\n1. " + ExcelReview.FeeReason + "，此项负偏差为10000.00元。", "全部明细低于门槛时仍有完整单项工程末条");
            state = V9Example(); balance = state.Balances.Single(); state.Items[1].Delta = 100; state.Items[1].AuditAmount = 1100; balance.NetOverride = 8500;
            text = ExcelReview.Compose(state, true); Assert(text.Contains("正偏差为400.00元") && V9TextDelta(text) == -8500, "核增核减混合按有符号净值配平并自动选择正偏差");
            balance.NetOverride = -1000; Assert(V9TextDelta(ExcelReview.Compose(state, true)) == 1000, "单项工程净核增支持负数人工总额");
            balance.NetOverride = 8900; Assert(!ExcelReview.Compose(state, true).Contains(ExcelReview.FeeReason), "差额为零时不生成零金额补差条款");
            balance.NetOverride = 0; Assert(ExcelReview.CalculateBalance(state, balance).TotalDelta == 0 && V9TextDelta(ExcelReview.Compose(state, true)) == 0, "显式零总额与未填写严格区分");
            balance.NetOverride = null; state.Items[1].Target = ExcelReview.Targets[1];
            Assert(ExcelReview.Targets.Sum(t => V9TextDelta(ExcelReview.Compose(state, true, t))) == -10000 && ExcelReview.Targets.Sum(t => Regex.Matches(ExcelReview.Compose(state, true, t), ExcelReview.FeeReason).Count) == 1, "同一工程跨正文归属只补一次并扣除各归属已列明细");
            var duplicateGroup = ExcelReview.Parse(new List<ExcelSheet> { V9Total("0_1_计算表", "同名工程", "5000", "4000"), ReviewTestSheet("0_1_明细", "同名工程", "墙", "200厚", "1000", "900"), V9Total("0_2_计算表", "同名工程", "7000", "5000"), ReviewTestSheet("0_2_明细", "同名工程", "梁", "C30", "1000", "800") }, "多专业.xlsx");
            Assert(duplicateGroup.Balances.Count == 1 && ExcelReview.CalculateBalance(duplicateGroup, duplicateGroup.Balances[0]).TotalDelta == -3000 && V9TextDelta(ExcelReview.Compose(duplicateGroup, true)) == -3000, "同名工程多张计算表分别对应后求和不丢失专业总额");
            var fallback = ExcelReview.Parse(new List<ExcelSheet> { V9Total("专业计算表", "甲楼", "1000", "800"), ReviewTestSheet("明细", "甲楼", "砌体", "200厚", "1000", "900") }, "无编号.xlsx");
            Assert(ExcelReview.BalanceIssues(fallback).Count == 0, "无编号表按前一张同工程计算表关联");
            var ambiguous = ExcelReview.Parse(new List<ExcelSheet> { V9Total("0_1_计算表甲", "甲楼", "1000", "800"), V9Total("0_1_计算表乙", "甲楼", "1000", "700"), ReviewTestSheet("0_1_明细", "甲楼", "砌体", "200厚", "1000", "900") }, "重复编号.xlsx");
            Reject(() => ExcelReview.Compose(ambiguous, true), "同一编号多个同名计算表时要求核对而不误加总");
            var conflict = V9Example(); conflict.Items[0].CalculationKey = "不对应的计算表";
            Reject(() => ExcelReview.Compose(conflict, true), "计算表明细未完整对应时阻止静默配平");
            Assert(ExcelReview.Compose(conflict, false).Contains("总额差额待核对"), "草稿明确显示总额对应问题");
            conflict.Balances[0].NetOverride = 12000; Assert(V9TextDelta(ExcelReview.Compose(conflict, true)) == -12000, "核实后填写人工总额可明确覆盖未关联来源");
            var renamed = V9Example(); renamed.Items[0].Group = "新工程"; Reject(() => ExcelReview.Compose(renamed, true), "明细改工程归属后禁止原总额继续静默补足造成重复计算");
            var noSource = ExcelReview.Parse(new List<ExcelSheet> { ReviewTestSheet("明细", "甲楼", "砌体", "200厚", "1000", "900") }, "仅明细.xlsx");
            Assert(!noSource.Balances[0].Enabled && !ExcelReview.Compose(noSource, true).Contains(ExcelReview.FeeReason), "仅有明细不臆造单项工程总额");
            noSource.Balances[0].Enabled = true; Reject(() => ExcelReview.Compose(noSource, true), "无计算表启用补差后要求提供总额");
            var nullSide = ExcelReview.ReadTotals(new List<ExcelSheet> { V9Total("计算表", "甲", "3258", "", "", "3258"), V9Total("计算表二", "乙", "", "-572394", "", "572394") });
            Assert(nullSide[0].Delta == -3258 && !nullSide[0].Audit.HasValue && nullSide[1].Delta == -572394 && !nullSide[1].Sent.HasValue, "总额一侧为空时采用明确核增减值且空金额仍为空");
            var formula = V9Total("计算表", "甲", "1000", "#公式未缓存");
            Assert(!ExcelReview.ReadTotals(new List<ExcelSheet> { formula })[0].Delta.HasValue, "总额公式无缓存不当作零");
            var partialCache = V9Total("计算表", "甲", "", "", "#公式未缓存", "200"); Assert(ExcelReview.ReadTotals(new List<ExcelSheet> { partialCache })[0].Message.Contains("未缓存公式"), "核增缺失公式缓存时不以零计算净差额");
            var disagree = V9Total("计算表", "甲", "1000", "800", "", "300");
            Assert(ExcelReview.ReadTotals(new List<ExcelSheet> { disagree })[0].Message.Contains("不一致"), "总额送审审核与显式核减冲突可定位");
            ExcelRow(formula, 4, "金额单位：万元"); Assert(ExcelReview.ReadTotals(new List<ExcelSheet> { formula })[0].Message.Contains("万元"), "万元总额不混入元口径");
            var repeated = V9Total("计算表", "甲", "1000", "800"); ExcelRow(repeated, 20, "", "建设工程造价", "", "1000", "800");
            Assert(ExcelReview.ReadTotals(new List<ExcelSheet> { repeated })[0].Message.Contains("多个"), "同表重复总额需核实不重复加总");

            string longFeature = "1、外墙砌块\n2、墙体厚度200mm\n3、砂浆强度等级Mb5.0\n4、砌块强度A5.0\n5、满足设计及规范要求，施工详图详见相关建筑施工图及工程清单所列全部工作内容\n6、未尽事宜按施工图纸及现行规范执行";
            var concise = ExcelReview.Parse(new List<ExcelSheet> { ReviewTestSheet("明细一", "甲楼", "砌块墙", longFeature, "1000", "900"), ReviewTestSheet("明细二", "甲楼", "砌块墙", longFeature.Replace("Mb5.0", "Mb7.5"), "100", "99") }, "特征精简.xlsx");
            concise.Threshold = 50; var descriptions = ExcelReview.BuildDescriptions(concise);
            Assert(descriptions[concise.Items[0]].Length < 40 && descriptions[concise.Items[0]].Contains("Mb5.0") && !descriptions[concise.Items[0]].Contains("200mm") && !descriptions[concise.Items[0]].Contains("未尽事宜"), "同名项目只提取最短区别特征而不重复共同厚度");
            Assert(descriptions[concise.Items[0]] != descriptions[concise.Items[1]] && concise.Items[0].Feature == longFeature, "精简不合并不同规格且不改动原始特征");
            concise.CompactFeatures = false; Assert(ExcelReview.Compose(concise, true).Contains("未尽事宜"), "关闭精简可恢复完整特征文字");
            concise.CompactFeatures = true; concise.Items[0].AuditName = "砌块墙"; concise.Items[0].AuditFeature = longFeature.Replace("200mm", "150mm");
            string paired = ExcelReview.BuildDescriptions(concise)[concise.Items[0]]; Assert(paired.Contains("送审为") && paired.Contains("审核为") && paired.Contains("200mm") && paired.Contains("150mm"), "精简后送审审核的不同规格仍分别交代");
            var legacy = new ExcelReviewState { ModelVersion = 8 }; ExcelReview.Normalize(legacy); Assert(legacy.CompactFeatures && legacy.Totals.Count == 0 && legacy.Balances.Count == 0, "旧项目启用特征精简且不凭空补历史总额");
            var layers = ExcelReview.Parse(new List<ExcelSheet> { ReviewTestSheet("明细", "甲楼", "抹灰", "内墙NQ1a\n1.乳胶漆两度；（另列）\n2.3厚腻子满墙刮2遍；（另列）\n3.10厚水泥砂浆罩面，压实赶光\n4.12厚水泥砂浆打底、扫光\n5.未尽事宜按设计图纸及规范要求执行", "1000", "900") }, "多层构造.xlsx");
            string layerText = ExcelReview.Compose(layers, true); Assert(layerText.Contains("抹灰，送审金额") && !layerText.Contains("水泥砂浆") && !layerText.Contains("3.10厚") && !layerText.Contains("2.3厚"), "名称无歧义时多层构造特征整体省略");

            state = V9Example(); state.Threshold = 500; balance = state.Balances[0];
            var p = new Project { Values = new Dictionary<string, string>(example.Values) }; p.Values["有关事项说明"] = ""; p.Values["项目名称"] = "一号楼差额验证";
            ExcelReview.StartLink(p, state); string saved = Path.Combine(output, "v090_配平项目.json"); Core.Save(saved, p);
            var loaded = Core.Load(saved); Assert(loaded.Excel.CompactFeatures && loaded.Excel.Totals[0].Source.Contains("D12") && loaded.Excel.Items[0].CalculationKey == "0_1_计算表" && V9TextDelta(ExcelReview.Compose(loaded.Excel, true)) == -10000, "项目保存重开保留来源、明细关联与配平设置");
            byte[] report = Core.Export(p, template, false); var reportDoc = XDocument.Parse(Read(report, "word/document.xml"));
            var tail = reportDoc.Descendants(Core.W + "p").Single(par => Core.RawParagraphText(par).Contains(ExcelReview.FeeReason));
            var donor = XDocument.Parse(Read(template, "word/document.xml")).Descendants(Core.W + "p").First(par => Core.RawParagraphText(par).Contains("{{基期价格}}"));
            Assert(tail.Element(Core.W + "pPr").Element(Core.W + "numPr") != null && XNode.DeepEquals(tail.Element(Core.W + "pPr").Element(Core.W + "spacing"), donor.Element(Core.W + "pPr").Element(Core.W + "spacing")), "末条补差沿用当前模板原生编号及有关事项行距");
            Assert(Core.Preview(p, template).Contains("2. " + ExcelReview.FeeReason), "正式Word预览包含连续编号末条");
            File.WriteAllBytes(Path.Combine(output, "v090_一万元配平验证.docx"), report);
            using (var form = new MainForm(true)) {
                V7Call(form, "LoadIntoForm", p);
                var compactBox = V7Field<CheckBox>(form, "excelConcise"); compactBox.Checked = false; Assert(!state.CompactFeatures, "界面精简选项修改同步数据"); compactBox.Checked = true;
                using (var dialog = (Form)V7Call(form, "CreateBalanceDialog")) {
                    V9Screenshot(dialog, Path.Combine(output, "v090_总额差额窗口.png")); var grid = ControlsBelow(dialog.Controls).OfType<DataGridView>().Single();
                    Assert(Convert.ToString(grid.Rows[0].Cells["Total"].Value) == "10000.00" && Convert.ToString(grid.Rows[0].Cells["After"].Value) == "10000.00", "差额窗口显示总额和报告核对值");
                    grid.Rows[0].Cells["Override"].Value = "10500.00"; grid.Rows[0].Cells["Reason"].Value = "经核实调整有关费用";
                    Assert(balance.NetOverride == 10500 && Core.Get(p, ExcelReview.Targets[0]).Contains("经核实调整有关费用，此项负偏差为1500.00元"), "人工总额及末条说明编辑立即联动正文");
                    grid.Rows[0].Cells["Override"].Value = ""; grid.Rows[0].Cells["Reason"].Value = ExcelReview.FeeReason;
                    Assert(!balance.NetOverride.HasValue && Convert.ToString(grid.Rows[0].Cells["Residual"].Value) == "1000.00", "清空人工总额恢复计算表总额和差额");
                    state.Totals[0].Message = "测试总额冲突"; grid.Rows[0].Cells["Enabled"].Value = false; grid.Rows[0].Cells["Enabled"].Value = true;
                    Assert(grid.Rows[0].DefaultCellStyle.BackColor == Color.MistyRose && Convert.ToString(grid.Rows[0].Cells["Check"].Value).Contains("测试总额冲突"), "总额冲突行颜色标识且说明具体原因");
                    V9Screenshot(dialog, Path.Combine(output, "v090_总额问题标识.png")); state.Totals[0].Message = ""; V7Call(form, "Changed"); V7Call(form, "UpdateExcelPreview");
                }
                form.Screenshot(Path.Combine(output, "v090_Excel新入口.png"), 4);
                V7Field<ComboBox>(form, "excelTarget").SelectedItem = ExcelReview.Targets[1]; V7Call(form, "AssignExcelTarget"); Assert(balance.Target == ExcelReview.Targets[1] && Core.Get(p, ExcelReview.Targets[1]).Contains(ExcelReview.FeeReason), "整工程改报告归属时末条随明细迁移");
                state.CompactFeatures = false; balance.NetOverride = 12345.67m; balance.Reason = "人工核实后的说明"; Core.Save(saved, p); loaded = Core.Load(saved);
                Assert(!loaded.Excel.CompactFeatures && loaded.Excel.Balances[0].NetOverride == 12345.67m && loaded.Excel.Balances[0].Reason == "人工核实后的说明", "人工总额说明及关闭精简状态保存重开不丢失");
            }
            string actualPath = @"C:\AIwork\明细对比表.xlsx";
            if (File.Exists(actualPath)) {
                byte[] before = System.Security.Cryptography.SHA256.Create().ComputeHash(File.ReadAllBytes(actualPath));
                var sheets = ExcelReview.Read(actualPath); var actual = ExcelReview.Parse(sheets, "明细对比表.xlsx"); actual.Threshold = 1000;
                Assert(sheets.Count == 40 && actual.Totals.Count == 13 && actual.Balances.Count == 3, "实际40张工作表识别13张计算表及3个单项工程");
                Assert(ExcelReview.BalanceIssues(actual).Count == 0 && actual.Items.All(i => !ExcelReview.IsCalculation(sheets.Single(s => s.Name == i.Sheet))), "实际计算表与明细全部对应且计算表不混入明细");
                var expected = new Dictionary<string, decimal> { { "1#、2#楼", 3074185 }, { "3#楼", 3258455 }, { "1#、2#、3#楼", 590424 } };
                foreach (var b in actual.Balances) { var r = ExcelReview.CalculateBalance(actual, b); Assert(-r.TotalDelta == expected[b.Group] && r.Listed + r.Residual == r.TotalDelta, "实际单项工程总额及配平：" + b.Group); }
                Assert(V9TextDelta(ExcelReview.Compose(actual, false)) == -6923064, "真实文件门槛一千元下报告偏差总和等于全部计算表总额");
                File.WriteAllText(Path.Combine(output, "v090_实际Excel组稿.txt"), ExcelReview.Compose(actual, false), Encoding.UTF8);
                File.WriteAllText(Path.Combine(output, "v090_实际总额核对.txt"), string.Join("\r\n\r\n", actual.Balances.Select(b => { var r = ExcelReview.CalculateBalance(actual, b); return b.Group + "\r\n" + r.Source + "\r\n总额净核减：" + ExcelReview.Money(-r.TotalDelta) + "；已列：" + ExcelReview.Money(-r.Listed) + "；末条：" + ExcelReview.Money(-r.Residual); })), Encoding.UTF8);
                actual.CompactFeatures = false; var fullDescriptions = ExcelReview.BuildDescriptions(actual); int fullLength = fullDescriptions.Sum(d => d.Value.Length);
                actual.CompactFeatures = true; var shortDescriptions = ExcelReview.BuildDescriptions(actual); int conciseLength = shortDescriptions.Sum(d => d.Value.Length);
                Assert(conciseLength < fullLength, "真实文件生成描述总长度得到缩短");
                File.WriteAllText(Path.Combine(output, "v090_实际文字精简对照.txt"), "全部明细描述字数：" + fullLength + " → " + conciseLength + "\r\n\r\n" + string.Join("\r\n\r\n", actual.Items.Where(i => fullDescriptions[i] != shortDescriptions[i]).Select(i => i.Location + "\r\n原文：" + fullDescriptions[i] + "\r\n精简：" + shortDescriptions[i])), Encoding.UTF8);
                var realProject = new Project { Values = new Dictionary<string, string>(example.Values) }; ExcelReview.StartLink(realProject, actual);
                using (var form = new MainForm(true)) { V7Call(form, "LoadIntoForm", realProject); using (var dialog = (Form)V7Call(form, "CreateBalanceDialog")) V9Screenshot(dialog, Path.Combine(output, "v090_实际工程总额.png")); }
                Assert(before.SequenceEqual(System.Security.Cryptography.SHA256.Create().ComputeHash(File.ReadAllBytes(actualPath))), "实际Excel验证前后内容完全一致");
            }
        }
    }
}
