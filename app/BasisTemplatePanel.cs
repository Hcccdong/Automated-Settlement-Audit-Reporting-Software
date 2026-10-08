using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SettlementReview {
    public partial class MainForm {
        string BasisLibraryPath { get { return Path.Combine(projectData, "basis-templates.json"); } }
        Form CreateBasisTemplateDialog(TextBox report) {
            var templates = BasisTemplateStore.Load(BasisLibraryPath);
            var dialog = new Form { Text = "审核依据模板管理", Font = new Font(Font.FontFamily, 9F), Size = new Size(900, 650), MinimumSize = new Size(650, 450), StartPosition = FormStartPosition.CenterParent };
            var chooser = new ComboBox { Dock = DockStyle.Top, DropDownStyle = ComboBoxStyle.DropDownList, Name = "BasisTemplateChoice" };
            var name = new TextBox { Dock = DockStyle.Top, Name = "BasisTemplateName", MaxLength = 150 };
            var content = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Both, Name = "BasisTemplateText", MaxLength = 20000 };
            var hint = new Label { Dock = DockStyle.Bottom, Height = 32, Text = "保存模板名称及内容后，可跨项目使用。请核对依据版本和适用时间。" };
            var actions = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 44 };
            string id = Guid.NewGuid().ToString("N");
            Action reload = () => { chooser.Items.Clear(); chooser.Items.AddRange(templates.Cast<object>().ToArray()); };
            chooser.SelectedIndexChanged += (s,e) => { var selected = chooser.SelectedItem as BasisTemplate; if (selected == null) return; id = selected.Id; name.Text = selected.Name; content.Text = selected.Text; };
            var create = new Button { Text = "新建模板", Width = 110, Height = 32 };
            create.Click += (s,e) => { chooser.SelectedIndex = -1; id = Guid.NewGuid().ToString("N"); name.Text = ""; content.Text = report.Text; name.Focus(); hint.Text = "填写新模板名称；内容可手动修改。"; };
            var save = new Button { Text = "保存名称与内容", Width = 150, Height = 32 };
            save.Click += (s,e) => { try { BasisTemplateStore.Save(BasisLibraryPath, templates, id, name.Text, content.Text); reload(); chooser.SelectedItem = templates.First(t => t.Id == id); hint.Text = "模板名称与内容已保存。"; } catch (Exception ex) { hint.Text = ex.Message; } };
            var use = new Button { Text = "使用当前内容", Width = 140, Height = 32 };
            use.Click += (s,e) => { if (string.IsNullOrWhiteSpace(content.Text)) { hint.Text = "请填写依据内容。"; return; } current.Values["当前依据选择内容"] = content.Text; report.Text = content.Text; dialog.DialogResult = DialogResult.OK; };
            actions.Controls.Add(create); actions.Controls.Add(save); actions.Controls.Add(use);
            content.Text = report.Text; reload();
            dialog.Controls.Add(content); dialog.Controls.Add(new Label { Dock = DockStyle.Top, Height = 28, Text = "依据内容（可编辑）" }); dialog.Controls.Add(name); dialog.Controls.Add(new Label { Dock = DockStyle.Top, Height = 28, Text = "模板名称（修改后点击保存即可改名）" }); dialog.Controls.Add(chooser); dialog.Controls.Add(new Label { Dock = DockStyle.Top, Height = 28, Text = "已保存的个人依据模板" }); dialog.Controls.Add(hint); dialog.Controls.Add(actions);
            return dialog;
        }
        void ManageBasisTemplates(TextBox report) { Run(() => { using (var dialog = CreateBasisTemplateDialog(report)) dialog.ShowDialog(this); }); }
    }
}
