using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace SettlementReview {
    public static partial class Tests {
        static int count;
        static void Assert(bool test, string reason) { if (!test) throw new Exception(reason); count++; }
        static void Reject(Action action, string reason) { bool rejected = false; try { action(); } catch (InvalidDataException) { rejected = true; } Assert(rejected, reason); }
        static string Read(byte[] zipBytes, string path) { using (var ms = new MemoryStream(zipBytes)) using (var z = new ZipArchive(ms, ZipArchiveMode.Read)) using (var r = new StreamReader(z.GetEntry(path).Open())) return r.ReadToEnd(); }
        static byte[] Change(byte[] bytes, string path, string xml) {
            using (var ms = new MemoryStream()) {
                ms.Write(bytes, 0, bytes.Length); ms.Position = 0;
                using (var z = new ZipArchive(ms, ZipArchiveMode.Update, true)) {
                    var e = z.GetEntry(path); if (e != null) e.Delete();
                    using (var w = new StreamWriter(z.CreateEntry(path).Open(), new UTF8Encoding(false))) w.Write(xml);
                } return ms.ToArray();
            }
        }
        public static int Run(string output) {
            Directory.CreateDirectory(output);
            var p = new Project(); foreach (var f in Core.Fields) p.Values[f.Name] = "无";
            p.Values["合同工期日历天数"] = "210"; p.Values["计划开工"] = "2021-11-15"; p.Values["计划竣工"] = "2022-06-12"; p.Values["实际开工"] = "2021-11-15"; p.Values["实际竣工"] = "2022-12-09"; p.Values["净核减金额"] = "80";
            p.Values["项目名称"] = "功能验证工程（仅供测试）"; p.Values["报告编号"] = "测试〔2026〕001号";
            p.Values["审核单位"] = "测试审核单位"; p.Values["委托单位"] = "测试委托单位";
            p.Values["报告日期"] = "2026年9月13日"; p.Values["合同金额"] = "1000";
            p.Values["金额比较口径"] = "送审金额"; p.Values["税费口径"] = "仅作测试，不代表实际税费口径";
            p.Values["送审金额"] = "1000"; p.Values["审定金额"] = "920"; p.Values["核减合计"] = "100"; p.Values["审增合计"] = "20";
            p.Values["审核依据及原则"] = "测试用依据第一行\r\n测试用依据第二行";
            p.Values["有关事项说明"] = "测试事项第一行\r\n测试事项第二行";
            p.Values["基期价格"] = "测试事项第一行\r\n测试事项第二行";
            var v = Core.Values(p, false); Assert(v["净核减金额"] == "80.00" && v["净核减率"] == "8.00%", "净额计算");
            p.Values["审定金额"] = "1020"; p.Values["核减合计"] = "0";
            p.Values["净核减金额"] = "-20";
            v = Core.Values(p, false); Assert(v["净核减金额"] == "-20.00" && v["净核减率"] == "-2.00%", "净核增计算");
            p.Values["送审金额"] = "0"; p.Values["审定金额"] = "0"; p.Values["审增合计"] = "0";
            p.Values["净核减金额"] = "0";
            Assert(Core.Values(p, false)["净核减率"].StartsWith("不适用"), "零分母");
            p.Values["送审金额"] = ""; p.Values["净核减金额"] = ""; Assert(Core.Values(p, true)["净核减金额"] == "待填写", "空值不是零");
            Reject(() => Core.Values(p, false), "正式稿不允许缺项");
            foreach (var invalid in new[] { "1万元", "1,000", "-1", "1.001", "1e3", "NaN", "", "9999999999999999" }) Reject(() => Core.Money(invalid, "测试金额"), "拒绝非法金额 " + invalid);
            p.Values["送审金额"] = "1000"; p.Values["审定金额"] = "920"; p.Values["核减合计"] = "99"; p.Values["审增合计"] = "20";
            p.Values["净核减金额"] = "80";
            Assert(Core.AmountCheck(p).Contains("提示：") && Core.Values(p, false)["净核减金额"] == "80.00", "差额不平衡只提示并保留用户输入"); p.Values["核减合计"] = "100";
            byte[] builtin = Core.BuiltIn(); var known = Core.ValidateTemplate(builtin); Assert(known.Contains("项目名称") && known.Contains("净核减金额") && known.Contains("合同标内调整说明"), "默认字段识别");
            byte[] report = Core.Export(p, builtin, false); string document = Read(report, "word/document.xml");
            Assert(!document.Contains("{{") && document.Contains("80.00"), "所有字段替换");
            Assert(document.Contains("<w:br") && document.Contains("测试事项第二行"), "多行正文保留");
            Assert(!document.Contains("温州体育") && !document.Contains("17983793"), "不得带入源样例事实");
            Assert(Read(builtin, "word/styles.xml") == Read(report, "word/styles.xml"), "样式保留");
            Assert(Read(builtin, "word/_rels/document.xml.rels") == Read(report, "word/_rels/document.xml.rels"), "关系保留");
            Assert(Core.Amounts.Length == 5 && string.Join("、", Core.Amounts.Select(f => f.Label)) == "送审金额、审核金额、核减金额、核增金额、净核减 / 净核增金额", "金额界面五项含净额输入");
            Assert(Core.RequiredForTemplate(new[] { "审核结论" }).Contains("净核减金额"), "整段结论模板也要求填写净核减金额");
            Assert(!Core.Fields.Any(f => new[] { "抄送单位", "审核内容", "审核程序及方法", "审核人员", "复核人员", "税费口径", "审核结论补充说明" }.Contains(f.Name)), "移除指定录入窗口");
            p.Values["建设单位"] = "建设方"; p.Values["施工单位"] = "施工方"; p.Values["代建单位"] = "代建方"; p.Values["抄送单位"] = "旧手填内容";
            Assert(Core.Values(p, false)["抄送单位"] == "建设方、施工方、代建方", "抄送顺序自动生成且不使用旧值");
            p.Values["代建单位"] = "建设方";
            Assert(Core.Values(p, false)["抄送单位"] == "建设方、施工方", "抄送去重");
            p.Values["代建单位"] = "无";
            Assert(Core.Values(p, false)["抄送单位"] == "建设方、施工方", "无代建时忽略");
            Assert(Core.Values(p, false)["审核依据及原则"] == Core.Get(p, "审核依据及原则"), "审核依据采用用户编辑内容");
            Assert(Core.Values(p, false)["审核结论"] == "本工程建设单位审定金额为1000.00元；审核金额为920.00元；核增金额为20.00元；核减金额为100.00元；净核减金额为80.00元。", "结论自动采用净核减方向");
            Assert(document.Contains("测试用依据第二行"), "导出采用手动依据正文");
            byte[] raw = Core.SourceTemplate(); var pristine = (byte[])raw.Clone(); byte[] ready = Core.PrepareTemplate(raw);
            Assert(raw.SequenceEqual(pristine), "导入预处理不改动来源");
            Assert(!Read(ready, "word/settings.xml").Contains("documentProtection") && !Read(ready, "word/settings.xml").Contains("trackRevisions"), "自动解除副本保护和修订记录");
            Assert(Core.ValidateTemplate(ready).Contains("净核减金额"), "无标记原报告自动识别");
            foreach (string part in new[] { "word/styles.xml", "word/numbering.xml", "word/fontTable.xml", "word/header1.xml", "word/footer1.xml", "word/footer2.xml", "word/_rels/document.xml.rels" }) Assert(Read(raw, part) == Read(ready, part), "原版部件保留 " + part);
            Assert(XNode.DeepEquals(XDocument.Parse(Read(raw, "word/document.xml")).Descendants(Core.W + "sectPr").Single(), XDocument.Parse(Read(ready, "word/document.xml")).Descendants(Core.W + "sectPr").Single()), "页尺寸边距及页眉页脚关系不变");
            var fixedOriginal = XDocument.Parse(Read(raw, "word/document.xml")).Descendants(Core.W + "p").Where(x => string.Concat(x.Descendants(Core.W + "t").Select(t => t.Value)).StartsWith("《浙江省建设工程计价规则》")).Single();
            var fixedPrepared = XDocument.Parse(Read(ready, "word/document.xml")).Descendants(Core.W + "p").Where(x => string.Concat(x.Descendants(Core.W + "t").Select(t => t.Value)).Contains("{{审核依据及原则}}")).Single();
            Assert(XNode.DeepEquals(fixedOriginal.Element(Core.W + "pPr"), fixedPrepared.Element(Core.W + "pPr")), "依据关联保留原段落格式");
            Assert(Read(ready, "word/settings.xml").Contains("updateFields"), "文档打开时刷新页码域");
            var signed = XDocument.Parse(Read(ready, "word/document.xml")).Descendants(Core.W + "p").Where(x => string.Concat(x.Descendants(Core.W + "t").Select(t => t.Value)).Contains("{{报告日期}}")).Single();
            Assert(string.Concat(signed.Descendants(Core.W + "t").Select(t => t.Value)).StartsWith("                                      "), "签署日期原有定位空格保留");
            byte[] revisions = Change(builtin, "word/header1.xml", "<w:hdr xmlns:w='" + Core.W + "'><w:p><w:del><w:r><w:delText>旧删除内容</w:delText></w:r></w:del><w:ins><w:r><w:t>保留新增内容</w:t></w:r></w:ins><w:moveFrom><w:r><w:t>移出旧内容</w:t></w:r></w:moveFrom><w:moveTo><w:r><w:t>保留移入内容</w:t></w:r></w:moveTo></w:p></w:hdr>");
            string cleaned = Read(Core.PrepareTemplate(revisions), "word/header1.xml");
            Assert(cleaned.Contains("保留新增内容") && cleaned.Contains("保留移入内容") && !cleaned.Contains("旧删除内容") && !cleaned.Contains("移出旧内容") && !cleaned.Contains("w:ins") && !cleaned.Contains("w:moveTo"), "自动接受修订并移除已删除内容");
            File.WriteAllBytes(Path.Combine(output, "原模板直接导出验证.docx"), Core.Export(p, raw, false));
            File.WriteAllBytes(Path.Combine(output, "主模板自动适配.docx"), ready);
            var imported = new Project(); Core.ApplyTemplateDefaults(imported, raw, ready);
            Assert(Core.Get(imported, "审核单位") == "温州市建设工程咨询有限公司" && Core.Get(imported, "建设单位") == "温州体育中心", "从模板回填单位信息");
            Assert(Core.Get(imported, "合同金额") == "20285176" && Core.Get(imported, "工程概况").Contains("场馆外立面"), "回填合同金额及工程概况");
            Assert(Core.MustEnter.All(k => Core.Get(imported, k) == ""), "项目名称编号日期及审核金额留给用户填写");
            Assert(Core.ParseDate(Core.Get(imported, "计划开工")) == new DateTime(2021,11,15) && Core.ParseDate(Core.Get(imported, "实际竣工")) == new DateTime(2022,12,9), "四个日期按模板回填");
            var scheduleDefaults = Core.ScheduleValues(imported, true);
            Assert(scheduleDefaults["计划工期日历天数"] == "210" && scheduleDefaults["实际工期日历天数"] == "390", "模板日期自动计算首尾均计日历天");
            Assert(Core.Get(imported, "工期补充说明").Contains("罚款190000元") && scheduleDefaults["工期情况"].Contains("实际工期为390日历天"), "保留模板工期句式及可修改补充说明");
            imported.Values["审核单位"] = "用户已修改单位"; Core.ApplyTemplateDefaults(imported, raw, ready);
            Assert(Core.Get(imported, "审核单位") == "用户已修改单位", "重新导入不覆盖已编辑内容");
            Assert(Core.CalendarDays(new DateTime(2024,2,28), new DateTime(2024,3,1), true) == 3, "闰年跨月日历天");
            Assert(Core.CalendarDays(new DateTime(2026,1,1), new DateTime(2026,1,1), true) == 1 && Core.CalendarDays(new DateTime(2026,1,1), new DateTime(2026,1,1), false) == 0, "同日起止及两种计日口径");
            Reject(() => Core.CalendarDays(new DateTime(2026,2,1), new DateTime(2026,1,1), true), "反向日期提示");
            imported.Values["实际开工"] = "2024-02-28"; imported.Values["实际竣工"] = "2024-03-01";
            Assert(Core.ScheduleValues(imported, true)["工期情况"].Contains("实际工期为3日历天"), "修改日期后自动更新工期段落");
            imported.Values["实际竣工"] = "2024-02-30"; Reject(() => Core.ScheduleValues(imported, true), "无效日期不接受");
            p.Values["净核减金额"] = "90";
            Assert(Core.AmountCheck(p).Contains("相差 10.00 元"), "手填净核减不符时显示差额");
            Assert(Core.Values(p, false)["净核减金额"] == "90.00", "不自动改写用户的净核减输入");
            Assert(string.Concat(XDocument.Parse(Read(Core.Export(p, raw, false), "word/document.xml")).Descendants(Core.W + "t").Select(x => x.Value)).Contains("净核减金额为90.00元"), "金额差额只提示仍可导出用户输入");
            p.Values["净核减金额"] = "80"; Assert(!Core.AmountCheck(p).Contains("提示："), "改正后验算提示恢复一致");
            var originalDoc = XDocument.Parse(Read(raw, "word/document.xml"));
            var readyDoc = XDocument.Parse(Read(ready, "word/document.xml"));
            var finalDoc = XDocument.Parse(Read(Core.Export(p, raw, false), "word/document.xml"));
            var originalShapes = originalDoc.Descendants().Where(e => e.Name.LocalName == "AlternateContent").ToList();
            var readyShapes = readyDoc.Descendants().Where(e => e.Name.LocalName == "AlternateContent").ToList();
            var finalShapes = finalDoc.Descendants().Where(e => e.Name.LocalName == "AlternateContent").ToList();
            Assert(originalShapes.Count > 0 && originalShapes.Count == readyShapes.Count && originalShapes.Zip(readyShapes, (a,b) => XNode.DeepEquals(a,b)).All(b => b), "模板编号下横线绘图对象完整保留");
            Assert(originalShapes.Count == finalShapes.Count && originalShapes.Zip(finalShapes, (a,b) => XNode.DeepEquals(a,b)).All(b => b), "最终导出横线对象不变");
            var introAfter = readyDoc.Descendants(Core.W + "p").First(x => Core.RawParagraphText(x).StartsWith("我公司接受"));
            var underlined = introAfter.Descendants(Core.W + "r").Where(r => r.Element(Core.W + "rPr") != null && r.Element(Core.W + "rPr").Element(Core.W + "u") != null).ToList();
            Assert(underlined.Count(r => string.Concat(r.Descendants(Core.W + "t").Select(t => t.Value)).Contains("{{委托单位}}")) == 2 && underlined.Any(r => string.Concat(r.Descendants(Core.W + "t").Select(t => t.Value)).Contains("{{项目名称}}")), "三处原文下划线作用于替换字段");
            var introFinal = finalDoc.Descendants(Core.W + "p").First(x => Core.RawParagraphText(x).StartsWith("我公司接受"));
            Assert(introFinal.Descendants(Core.W + "r").Where(r => r.Element(Core.W + "rPr") != null && r.Element(Core.W + "rPr").Element(Core.W + "u") != null).Count(r => string.Concat(r.Descendants(Core.W + "t").Select(t => t.Value)).Contains("测试委托单位")) == 2, "输出委托单位下划线可见文字保留");
            Assert(Core.Body.Where(f => f.Name != "审核依据及原则").Take(5).Select(f => f.Name).SequenceEqual(Core.MatterFields) && !Core.Body.Any(f => f.Name == "有关事项说明"), "有关事项拆为五项");
            Assert(Core.MatterFields.All(k => Core.ValidateTemplate(ready).Contains(k)), "模板五个事项段落各自关联");
            Assert(Core.TemplateParagraphs(ready).Any(x => x.Text.Contains("工程下浮率{{工程下浮率}}；")), "下浮率保留模板句式与标点");
            var alternateDoc = XDocument.Parse(Read(raw, "word/document.xml"));
            foreach (var paragraph in alternateDoc.Descendants(Core.W + "p")) {
                string t = Core.RawParagraphText(paragraph); string old = "本工程建设单位审定金额为";
                if (t.Contains(old)) Core.ReplaceSpan(paragraph, t.IndexOf(old), old.Length, "本次送审金额为");
                old = "工程量、联系单、取费标准、材料价格等其他与造价相关的内容。";
                if (t.Contains(old)) Core.ReplaceSpan(paragraph, t.IndexOf(old), old.Length, "按本次委托范围开展专项审核。");
            }
            byte[] alternateSource = Change(raw, "word/document.xml", alternateDoc.ToString());
            string alternate = Read(Core.Export(p, alternateSource, false), "word/document.xml");
            string alternateText = string.Concat(XDocument.Parse(alternate).Descendants(Core.W + "t").Select(t => t.Value));
            Assert(alternateText.Contains("本次送审金额为1000.00元") && alternateText.Contains("按本次委托范围开展专项审核。"), "改换模板后采用该模板结论及固定正文");
            Assert(Core.ConclusionPreview(p, alternateSource).StartsWith("本次送审金额为1000.00元"), "金额页结论预览跟随当前模板");
            string simpleXml = "<w:document xmlns:w='" + Core.W + "'><w:body><w:p><w:r><w:t>自定义委托报告</w:t></w:r></w:p><w:p><w:r><w:t>工程：</w:t></w:r><w:r><w:rPr><w:u w:val='single'/><w:b/></w:rPr><w:t>旧工程名称</w:t></w:r><w:r><w:t>，固定说明保留。</w:t></w:r></w:p><w:sectPr/></w:body></w:document>";
            byte[] generic = Change(raw, "word/document.xml", simpleXml); byte[] genericReady = Core.PrepareTemplate(generic);
            Assert(Read(genericReady, "word/document.xml") == simpleXml && !Core.HasTemplateFields(genericReady), "无法识别的模板保持原样等待用户关联");
            var slot = Core.TemplateParagraphs(genericReady).First(x => x.Text.Contains("旧工程名称"));
            byte[] mapped = Core.MapSelection(genericReady, slot, slot.Text.IndexOf("旧工程名称"), "旧工程名称".Length, "项目名称");
            var projectOnly = new Project(); projectOnly.Values["项目名称"] = "新工程名称";
            string genericExport = Read(Core.Export(projectOnly, mapped, false), "word/document.xml");
            Assert(genericExport.Contains("新工程名称") && genericExport.Contains("固定说明保留。") && !genericExport.Contains("旧工程名称"), "通用模板关联后正确填写且不强制无关字段");
            var gdoc = XDocument.Parse(genericExport); var grun = gdoc.Descendants(Core.W + "r").First(r => r.Descendants(Core.W + "t").Any(t => t.Value == "新工程名称"));
            Assert(grun.Element(Core.W + "rPr").Element(Core.W + "u") != null && grun.Element(Core.W + "rPr").Element(Core.W + "b") != null, "手动关联保留下划线与加粗");
            File.WriteAllBytes(Path.Combine(output, "通用模板关联验证.docx"), Core.Export(projectOnly, mapped, false));
            using (var mappingForm = new TemplateMappingForm(genericReady)) {
                mappingForm.Show(); System.Windows.Forms.Application.DoEvents();
                var selector = mappingForm.Controls.Cast<System.Windows.Forms.Control>().SelectMany(c => c.Controls.Cast<System.Windows.Forms.Control>()).OfType<System.Windows.Forms.ComboBox>().Single();
                Assert(selector.SelectedItem != null && selector.SelectedItem.ToString() == "项目名称", "关联窗口默认填写项可用");
                using (var bitmap = new System.Drawing.Bitmap(mappingForm.Width, mappingForm.Height)) { mappingForm.DrawToBitmap(bitmap, new System.Drawing.Rectangle(0, 0, mappingForm.Width, mappingForm.Height)); bitmap.Save(Path.Combine(output, "界面_模板关联.png")); }
                mappingForm.Close();
            }
            var xml = XDocument.Parse(Read(builtin, "word/document.xml")); var body = xml.Root.Element(Core.W + "body");
            var split = new XElement(Core.W + "p",
                new XElement(Core.W + "r", new XElement(Core.W + "t", "前缀{{项目")),
                new XElement(Core.W + "r", new XElement(Core.W + "rPr", new XElement(Core.W + "b")), new XElement(Core.W + "t", "名称}}和{{项目名称}}后缀")));
            body.AddFirst(split);
            byte[] custom = Change(builtin, "word/document.xml", xml.ToString());
            custom = Change(custom, "word/header1.xml", "<w:hdr xmlns:w='" + Core.W + "'><w:p><w:r><w:t>{{项目</w:t></w:r><w:r><w:t>名称}}</w:t></w:r></w:p></w:hdr>");
            custom = Change(custom, "customXml/item1.xml", "<preserve>opaque data</preserve>");
            var changed = Core.Export(p, custom, false); string replaced = Read(changed, "word/document.xml");
            Assert(!replaced.Contains("{{") && replaced.Contains("后缀"), "跨run及重复字段");
            Assert(Read(changed, "word/header1.xml").Contains(p.Values["项目名称"]) && !Read(changed, "word/header1.xml").Contains("{{"), "页眉字段");
            Assert(Read(changed, "customXml/item1.xml") == Read(custom, "customXml/item1.xml"), "不透明部件保留");
            byte[] bad = Change(builtin, "word/document.xml", Read(builtin, "word/document.xml").Replace("{{项目名称}}", "{{未知字段}}"));
            Reject(() => Core.ValidateTemplate(bad), "拒绝未知字段");
            bad = Change(builtin, "word/document.xml", Read(builtin, "word/document.xml").Replace("{{项目名称}}", "{{项目名称}"));
            Reject(() => Core.ValidateTemplate(bad), "拒绝未闭合字段");
            bad = Change(builtin, "word/comments.xml", "<w:comments xmlns:w='" + Core.W + "'><w:t>{{项目名称}}</w:t></w:comments>");
            Reject(() => Core.ValidateTemplate(bad), "拒绝不支持位置的字段");
            p.Values["基期价格"] = "{{未替换的内容}}";
            Reject(() => Core.Export(p, builtin, false), "拒绝正文残留标记"); p.Values["基期价格"] = "无"; p.Values["有关事项说明"] = "无";
            Assert(Core.ResolveTemplate(Path.Combine("data", "templates", "example.docx")) == Path.Combine(Core.Root, "data", "templates", "example.docx"), "模板相对路径支持迁移");
            string projectPath = Path.Combine(output, "roundtrip.json"); Core.Save(projectPath, p); var loaded = Core.Load(projectPath);
            Assert(loaded.Id == p.Id && loaded.Values["审核依据及原则"] == p.Values["审核依据及原则"], "项目保存重开");
            p.Values["有关事项说明"] = "测试后续版本"; Core.Save(projectPath, p);
            Assert(File.Exists(projectPath + ".bak") && Core.Load(projectPath + ".bak").Values["有关事项说明"] == "无", "保存备份保留前一版");
            File.WriteAllBytes(Path.Combine(output, "测试审核稿.docx"), report);
            File.WriteAllBytes(Path.Combine(output, "测试草稿.docx"), Core.Export(new Project(), builtin, true));
            Assert(Read(Core.Export(new Project(), builtin, true), "word/document.xml").Contains("草稿  待审核确认"), "草稿标识");
            var excelExample = CheckExcel(output, p, builtin);
            excelExample = CheckVersionSix(output, p, builtin);
            CheckVersionSeven(output, p, builtin);
            CheckVersionEight(output, p, builtin);
            CheckVersionNine(output, p, builtin);
            CheckVersionTen(output, p, builtin);
            CheckVersionEleven(output, p, builtin);
            CheckVersionTwelve(output, p, builtin);
            CheckVersionThirteen(output);
            CheckVersionFourteen(output);
            CheckVersionFifteen(output);
            using (var form = new MainForm(true)) {
                form.Screenshot(Path.Combine(output, "界面_导入模板.png"), 0);
                var example = new Project(); Core.ApplyTemplateDefaults(example, raw, ready);
                typeof(MainForm).GetMethod("LoadIntoForm", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance).Invoke(form, new object[] { example });
                form.Screenshot(Path.Combine(output, "界面_基本信息.png"), 1);
                form.Screenshot(Path.Combine(output, "界面_工期情况.png"), 2);
                form.Screenshot(Path.Combine(output, "界面_审核正文.png"), 3);
                example.Values["送审金额"] = "1000"; example.Values["审定金额"] = "920"; example.Values["核减合计"] = "100"; example.Values["审增合计"] = "20"; example.Values["净核减金额"] = "90";
                example.Excel = excelExample;
                typeof(MainForm).GetMethod("LoadIntoForm", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance).Invoke(form, new object[] { example });
                form.Screenshot(Path.Combine(output, "界面_Excel对比.png"), 4);
                form.Show(); System.Windows.Forms.Application.DoEvents();
                typeof(MainForm).GetMethod("ShowLinkedBody", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance).Invoke(form, null);
                form.Screenshot(Path.Combine(output, "界面_联动正文.png"), 3);
                form.Screenshot(Path.Combine(output, "界面_金额核对.png"), 5);
                form.Screenshot(Path.Combine(output, "界面_模板导出.png"), 6);
                Assert(!ControlsBelow(form.Controls).OfType<System.Windows.Forms.TabControl>().First().TabPages[6].Controls.Cast<System.Windows.Forms.Control>().SelectMany(c => ControlsBelow(c.Controls).Concat(new[] { c })).OfType<System.Windows.Forms.Button>().Any(b => b.Text.Contains("导入")), "导出页没有重复导入按钮");
                form.Size = new System.Drawing.Size((int)(1060 * form.DeviceDpi / 96F), (int)(720 * form.DeviceDpi / 96F));
                form.Screenshot(Path.Combine(output, "界面_小窗口.png"), 2);
                form.Screenshot(Path.Combine(output, "界面_Excel小窗口.png"), 4);
                var compactGrid = V7Field<System.Windows.Forms.DataGridView>(form, "excelGrid"); Assert(compactGrid.ClientSize.Height >= compactGrid.ColumnHeadersHeight + compactGrid.Rows[0].Height * 2, "小窗口下Excel至少能显示两行明细（" + compactGrid.ClientSize.Height + "/" + compactGrid.ColumnHeadersHeight + "/" + compactGrid.Rows[0].Height + "）");
            }
            File.WriteAllText(Path.Combine(output, "测试结果.txt"), "通过 " + count + " 项检查。\r\n包括Word样式预览、精简模式按名称歧义提取最少特征、金额方向联动、四舍五入配平、项目保存及既有功能回归。\r\n真实Excel验证按文件是否存在执行，详情见实际总额核对及组稿材料。\r\n预览保留常用文字和段落格式；复杂绘图、页眉页脚及精确分页仍需Word或WPS复核。", Encoding.UTF8);
            return 0;
        }
    }
}
