using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace SettlementReview {
    [DataContract] public class BasisTemplate {
        [DataMember] public string Id = Guid.NewGuid().ToString("N");
        [DataMember] public string Name = "";
        [DataMember] public string Text = "";
        public override string ToString() { return Name; }
    }
    public static class BasisTemplateStore {
        public static List<BasisTemplate> Load(string path) {
            if (!File.Exists(path)) return new List<BasisTemplate>();
            using (var file = File.OpenRead(path)) {
                if (file.Length > 4 * 1024 * 1024) throw new InvalidDataException("依据模板库超过4MB，请核对文件。");
                var list = (List<BasisTemplate>)new DataContractJsonSerializer(typeof(List<BasisTemplate>)).ReadObject(file);
                if (list == null || list.Any(t => t == null || string.IsNullOrWhiteSpace(t.Id) || string.IsNullOrWhiteSpace(t.Name))) throw new InvalidDataException("依据模板库结构无效。");
                return list;
            }
        }
        public static void Save(string path, List<BasisTemplate> templates, string id, string name, string text) {
            name = (name ?? "").Trim(); text = (text ?? "").Trim();
            if (name == "" || text == "") throw new InvalidDataException("请填写模板名称和审核依据内容。");
            if (templates.Any(t => t.Id != id && string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase))) throw new InvalidDataException("已有同名依据模板，请使用其他名称。");
            var next = templates.Select(t => new BasisTemplate { Id = t.Id, Name = t.Name, Text = t.Text }).ToList();
            var item = next.FirstOrDefault(t => t.Id == id); if (item == null) { item = new BasisTemplate { Id = id }; next.Add(item); }
            item.Name = name; item.Text = text;
            using (var memory = new MemoryStream()) {
                new DataContractJsonSerializer(typeof(List<BasisTemplate>)).WriteObject(memory, next);
                if (memory.Length > 4 * 1024 * 1024) throw new InvalidDataException("依据模板库超过4MB，未保存。");
                ProjectStore.Atomic(path, memory.ToArray());
            }
            templates.Clear(); templates.AddRange(next);
        }
    }
}
