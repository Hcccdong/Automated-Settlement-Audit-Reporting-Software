using System;
using System.IO;
using System.Runtime.Serialization.Json;

namespace SettlementReview {
    // File-system boundary. No UI or report-generation dependencies.
    public static class ProjectStore {
        const int MaxProjectBytes = 16 * 1024 * 1024;
        const int MaxGraphItems = 1000000;
        public static void Save(string path, Project p) {
            p.Updated = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            using (var ms = new MemoryStream()) {
                new DataContractJsonSerializer(typeof(Project), new DataContractJsonSerializerSettings { MaxItemsInObjectGraph = MaxGraphItems }).WriteObject(ms, p);
                if (ms.Length > MaxProjectBytes) throw new InvalidDataException("项目内容超过16MB，请按单位工程分成项目副本保存；原保存文件保留。");
                Atomic(path, ms.ToArray());
            }
        }
        public static Project Load(string path) {
            using (var s = File.OpenRead(path)) {
                if (s.Length > MaxProjectBytes) throw new InvalidDataException("项目文件过大。");
                var p = (Project)new DataContractJsonSerializer(typeof(Project), new DataContractJsonSerializerSettings { MaxItemsInObjectGraph = MaxGraphItems }).ReadObject(s);
                Guid id; if (p == null || p.Values == null || !Guid.TryParseExact(p.Id, "N", out id)) throw new InvalidDataException("项目文件结构无效。");
                return p;
            }
        }
        public static void Atomic(string path, byte[] bytes) {
            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path)));
            string tmp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try {
                File.WriteAllBytes(tmp, bytes);
                if (File.Exists(path)) File.Replace(tmp, path, path + ".bak"); else File.Move(tmp, path);
            } finally { if (File.Exists(tmp)) File.Delete(tmp); }
        }
    }
}
