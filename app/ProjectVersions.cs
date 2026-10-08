using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace SettlementReview {
    public class ProjectGroup {
        public List<Project> Versions;
        public Project Latest { get { return Versions[0]; } }
        public override string ToString() { string name = Core.Get(Latest, "项目名称"); return name == "" ? "未命名项目" : name; }
    }
    public class ProjectVersion {
        public string Path, Label;
        public DateTime Saved;
        public override string ToString() { return Saved.ToString("yyyy-MM-dd HH:mm:ss.fff") + "  ·  " + Label; }
    }
    public static partial class Core {
        public static List<ProjectGroup> GroupProjects(IEnumerable<Project> projects, string activeId) {
            return projects.GroupBy(p => Get(p, "项目名称") == "" ? "id:" + p.Id : "name:" + Get(p, "项目名称"), StringComparer.CurrentCultureIgnoreCase)
                .Select(g => new ProjectGroup { Versions = g.OrderByDescending(p => p.LocalSavedAt).ThenByDescending(p => p.Updated).ThenByDescending(p => p.Id == activeId).ToList() }).OrderByDescending(g => g.Latest.LocalSavedAt).ThenByDescending(g => g.Latest.Updated).ToList();
        }
    }
    public partial class MainForm {
        List<ProjectVersion> ProjectVersionFiles(ProjectGroup group) {
            var versions = new List<ProjectVersion>();
            foreach (var project in group.Versions) {
                string path = System.IO.Path.Combine(projectDir, project.Id + ".json");
                if (File.Exists(path)) versions.Add(new ProjectVersion { Path = path, Label = "已保存版本", Saved = File.GetLastWriteTime(path) });
                string copies = System.IO.Path.Combine(projectData, "copies", project.Id);
                if (Directory.Exists(copies)) foreach (string copy in Directory.GetFiles(copies, "*.json")) versions.Add(new ProjectVersion { Path = copy, Label = "审核副本", Saved = File.GetLastWriteTime(copy) });
            }
            return versions.OrderByDescending(v => v.Saved).ToList();
        }
        void OpenProjectVersion(ProjectVersion version) {
            var p = Core.Load(version.Path);
            if (!LeaveCurrent()) return;
            // Opening a snapshot is a working state; the snapshot itself is never edited.
            LoadIntoForm(p); dirty = true; status.Text = "已打开所选审核版本，修改后点击保存项目；历史副本保留。";
        }
        void ShowProjectVersions() { Run(() => {
            var group = projects.SelectedItem as ProjectGroup;
            if (group == null) group = Core.GroupProjects(catalog, current.Id).FirstOrDefault(g => g.Versions.Any(p => p.Id == current.Id));
            if (group == null) { status.Text = "请先保存或选择一个工程。"; return; }
            using (var dialog = CreateProjectVersionDialog(group)) dialog.ShowDialog(this);
        }); }
        Form CreateProjectVersionDialog(ProjectGroup group) {
            var versions = ProjectVersionFiles(group);
            var dialog = new Form { Text = "审核版本 · " + group, Font = Font, AutoScaleMode = AutoScaleMode.Dpi, AutoScaleDimensions = new SizeF(96, 96), Size = new Size(780, 500), StartPosition = FormStartPosition.CenterParent, MinimizeBox = false };
                var list = new ListBox { Dock = DockStyle.Fill, HorizontalScrollbar = true }; list.Items.AddRange(versions.Cast<object>().ToArray()); if (list.Items.Count > 0) list.SelectedIndex = 0;
                var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 60, Padding = new Padding(12) };
                var open = Button("打开所选版本", (s,e) => Run(() => { if (list.SelectedItem == null) return; OpenProjectVersion((ProjectVersion)list.SelectedItem); dialog.Close(); }), 170); buttons.Controls.Add(open);
                buttons.Controls.Add(Button("关闭", (s,e) => dialog.Close(), 100));
                dialog.Controls.Add(list); dialog.Controls.Add(buttons); dialog.Controls.Add(new Label { Dock = DockStyle.Top, Height = 58, Padding = new Padding(12), Text = "同一工程的保存记录及审核副本集中显示。打开副本不会修改原副本文件。" });
            return dialog;
        }
    }
}
