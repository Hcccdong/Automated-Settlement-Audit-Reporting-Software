using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace SettlementReview {
    [DataContract] public class Project {
        [DataMember] public string Id = Guid.NewGuid().ToString("N");
        [DataMember] public Dictionary<string, string> Values = new Dictionary<string, string>();
        [DataMember] public string Template = "";
        [DataMember] public string Updated = "";
        [DataMember(EmitDefaultValue = false)] public ExcelReviewState Excel;
        public DateTime LocalSavedAt;
        public override string ToString() { return (Values.ContainsKey("项目名称") && Values["项目名称"] != "" ? Values["项目名称"] : "未命名项目") + "  ·  " + Updated; }
    }
    [DataContract] public class UiPreferences {
        [DataMember] public string LastExportDirectory = "";
    }
    public class Field {
        public string Name, Hint, Label; public int Lines;
        public Field(string name, string hint = "", int lines = 1, string label = null) { Name = name; Hint = hint; Lines = lines; Label = label ?? name; }
    }
}
