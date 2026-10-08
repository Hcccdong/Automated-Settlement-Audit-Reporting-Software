using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;

namespace SettlementReview {
    public static partial class Tests {
        static void CheckVersionThirteen(string output) {
            var itemA = new ReviewItem { Sheet = "明细建筑", Group = "原建筑", CalculationKey = "a", Delta = -10m, Include = true, Target = ExcelReview.Targets[0], Category = ExcelReview.Categories[0], Reason = ExcelReview.Reasons[0], Name = "测试甲", Warning = "", Note = "" };
            var itemB = new ReviewItem { Sheet = "明细安装", Group = "原安装", CalculationKey = "b", Delta = -20m, Include = true, Target = ExcelReview.Targets[0], Category = ExcelReview.Categories[0], Reason = ExcelReview.Reasons[0], Name = "测试乙", Warning = "", Note = "" };
            var state = new ExcelReviewState { ModelVersion = 11, Items = new List<ReviewItem> { itemA, itemB }, Totals = new List<ReviewTotal> { new ReviewTotal { Group = "原建筑", Key = "a", Delta = -12m, Source = "表A", Message = "", DetailSheets = new List<string> { "表A" } }, new ReviewTotal { Group = "原安装", Key = "b", Delta = -23m, Source = "表B", Message = "", DetailSheets = new List<string> { "表B" } } }, Balances = new List<ReviewBalance> { new ReviewBalance { Group = "原建筑", Target = ExcelReview.Targets[0], Enabled = true, Reason = ExcelReview.FeeReason }, new ReviewBalance { Group = "原安装", Target = ExcelReview.Targets[0], Enabled = true, Reason = ExcelReview.FeeReason } }, AmountDecimals = 2 };
            ReviewGroupEditor.Rename(state, new[] { itemA, itemB }, "合并工程");
            Assert(state.Items.All(i => i.Group == "合并工程") && state.Totals.All(t => t.Group == "合并工程") && state.Balances.Count == 1, "批量名称合并同步总额与唯一补差");
            var balance = ExcelReview.CalculateBalance(state, state.Balances[0]);
            Assert(balance.Issue == "" && balance.TotalDelta == -35m && balance.Residual == -5m, "名称合并保留总额与显示精度配平");
            Reject(() => ReviewGroupEditor.Rename(state, new[] { itemA }, " "), "批量改名拒绝空名称");
            Reject(() => ReviewGroupEditor.Rename(state, new ReviewItem[0], "新名称"), "批量改名拒绝未选择行");
            ReviewGroupEditor.Rename(state, new[] { itemA }, "拆分工程");
            Assert(itemB.Group == "合并工程" && itemA.Group == "拆分工程" && itemB.Delta == -20m, "仅修改选中行且不更改金额");
            Assert(state.Totals[0].Group == "拆分工程" && ExcelReview.CalculateBalance(state, state.Balances[0]).TotalDelta == -23m, "整张计算表关联明细迁移时总额随名称归属更新");

            var p = new Project(); p.Values["代建单位"] = "无"; p.Values["监理单位"] = "无"; p.Values["批复总投资"] = "无"; p.Values["建设单位"] = "保留的建设单位";
            string body = "<w:document xmlns:w='" + Core.W + "'><w:body><w:p><w:r><w:t>代建单位：{{代建单位}}</w:t></w:r></w:p><w:p><w:r><w:t>监理单位：{{监理单位}}</w:t></w:r></w:p><w:p><w:r><w:t>批复总投资：{{批复总投资}}</w:t></w:r></w:p><w:p><w:r><w:t>建设单位：{{建设单位}}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>";
            var template = Change(Core.BuiltIn(), "word/document.xml", body); var pristine = (byte[])template.Clone();
            string exported = Read(Core.Export(p, template, false), "word/document.xml");
            Assert(!exported.Contains("代建单位") && !exported.Contains("监理单位") && !exported.Contains("批复总投资") && exported.Contains("保留的建设单位"), "无项目信息整行删除且相邻有效行保留");
            Assert(template.SequenceEqual(pristine), "删行不改变原模板字节");
            p.Values["代建单位"] = "实际代建单位"; exported = Read(Core.Export(p, template, false), "word/document.xml");
            Assert(exported.Contains("实际代建单位"), "重新填写信息后导出恢复行");
            p.Values["审核依据及原则"] = "本项目手动审核依据\r\n合同约定资料";
            var basisTemplate = Change(template, "word/document.xml", body.Replace("<w:sectPr/>", "<w:p><w:r><w:t>{{审核依据及原则}}</w:t></w:r></w:p><w:sectPr/>"));
            Assert(Read(Core.Export(p, basisTemplate, false), "word/document.xml").Contains("本项目手动审核依据"), "编辑依据实际写入Word");
            string path = Path.Combine(output, "v013_审核依据保存.json"); Core.Save(path, p);
            Assert(Core.Get(Core.Load(path), "审核依据及原则") == p.Values["审核依据及原则"], "手动依据保存重开保持");
            Assert(Core.TemplateBasis(Core.SourceTemplate()).Contains("浙江省"), "依据选择来源读取实际模板");
            string legacyXml = "<w:document xmlns:w='" + Core.W + "'><w:body><w:p><w:r><w:t>{{项目名称}}</w:t></w:r></w:p><w:p><w:r><w:t>二、审核依据及原则</w:t></w:r></w:p><w:p><w:r><w:t>旧固定依据条目</w:t></w:r></w:p><w:p><w:r><w:t>三、审核内容</w:t></w:r></w:p><w:sectPr/></w:body></w:document>";
            var legacyTemplate = Change(template, "word/document.xml", legacyXml); p.Values["项目名称"] = "兼容工程";
            string legacyExport = Read(Core.Export(p, legacyTemplate, false), "word/document.xml");
            Assert(legacyExport.Contains("本项目手动审核依据") && !legacyExport.Contains("旧固定依据条目"), "旧版已关联模板的固定依据升级为可编辑位置");
            var basisDoc = XDocument.Parse(Read(Core.Export(p, basisTemplate, false), "word/document.xml"));
            Assert(basisDoc.Descendants(Core.W + "p").Count(x => Core.RawParagraphText(x) == "本项目手动审核依据" || Core.RawParagraphText(x) == "合同约定资料") == 2, "审核依据按条目输出独立段落");
            var tableDoc = XDocument.Parse("<w:document xmlns:w='" + Core.W + "'><w:body><w:tbl><w:tr><w:tc><w:p><w:r><w:t>监理</w:t></w:r><w:r><w:t>单位：</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>{{监理单位}}</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>其他信息保留</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>");
            Core.RemoveAbsentProjectLines(tableDoc, p);
            Assert(tableDoc.Descendants(Core.W + "tr").Count() == 1 && tableDoc.ToString().Contains("其他信息保留"), "无信息表格行删除支持标签拆分文本且保留相邻行");
            state.Balances[0].NetOverride = 23m;
            state.Balances.Add(new ReviewBalance { Group = "拆分工程", Target = ExcelReview.Targets[0], Enabled = true, NetOverride = 12m });
            Reject(() => ReviewGroupEditor.Rename(state, new[] { itemA }, "合并工程"), "人工总额冲突合并必须拒绝");
            Assert(itemA.Group == "拆分工程" && state.Totals[0].Group == "拆分工程" && state.Balances.Count == 2, "拒绝合并时原名称与总额设置保持");
            using (var form = new MainForm(true)) {
                var project = new Project { Excel = state }; V7Call(form, "LoadIntoForm", project);
                var inputs = V7Field<Dictionary<string, TextBox>>(form, "inputs");
                Assert(inputs.ContainsKey("审核依据及原则") && !inputs["审核依据及原则"].ReadOnly, "审核依据录入框可手动编辑");
                var choices = V7Field<ComboBox>(form, "batchGroup"); V7Call(form, "RefreshGroupChoices");
                Assert(choices.DropDownStyle == ComboBoxStyle.DropDown && choices.Items.Contains("拆分工程"), "批量单项名称支持自由输入与已有名称选择");
                Assert(!ControlsBelow(form.Controls).OfType<Button>().Any(b => b.Text == "预览文字"), "取消重复预览文字按钮");
                form.Screenshot(Path.Combine(output, "v013_批量单项名称.png"), 4);
                form.Screenshot(Path.Combine(output, "v013_审核依据编辑.png"), 3);
            }
            File.WriteAllBytes(Path.Combine(output, "v013_无信息删行与依据编辑.docx"), Core.Export(p, basisTemplate, false));
        }
    }
}
