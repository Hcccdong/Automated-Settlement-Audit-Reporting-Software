using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace SettlementReview {
    [DataContract] public class ReviewItem {
        [DataMember] public bool Include { get; set; }
        [DataMember] public bool Confirmed { get; set; }
        [DataMember] public string Sheet { get; set; }
        [DataMember] public int Row { get; set; }
        [DataMember] public string Group { get; set; }
        [DataMember] public string Category { get; set; }
        [DataMember] public string Target { get; set; }
        [DataMember] public string Code { get; set; }
        [DataMember] public string Name { get; set; }
        [DataMember] public string Feature { get; set; }
        [DataMember] public string Unit { get; set; }
        [DataMember] public string AuditCode { get; set; }
        [DataMember] public string AuditName { get; set; }
        [DataMember] public string AuditFeature { get; set; }
        [DataMember] public string AuditUnit { get; set; }
        [DataMember] public decimal? SentQuantity { get; set; }
        [DataMember] public decimal? AuditQuantity { get; set; }
        [DataMember] public decimal? SentPrice { get; set; }
        [DataMember] public decimal? AuditPrice { get; set; }
        [DataMember] public decimal? SentAmount { get; set; }
        [DataMember] public decimal? AuditAmount { get; set; }
        [DataMember] public decimal? SourceIncrease { get; set; }
        [DataMember] public decimal? SourceDecrease { get; set; }
        [DataMember] public decimal? Delta { get; set; }
        [DataMember] public string Reason { get; set; }
        [DataMember] public string Note { get; set; }
        [DataMember] public string Warning { get; set; }
        [DataMember] public string CalculationKey { get; set; }
        public string Direction { get { return !Delta.HasValue ? "待确认" : Delta > 0 ? "正偏差 / 核增" : Delta < 0 ? "负偏差 / 核减" : "无偏差"; } }
        public decimal? IncreaseAmount { get { return Delta.HasValue ? (decimal?)Math.Max(0, Delta.Value) : null; } }
        public decimal? DecreaseAmount { get { return Delta.HasValue ? (decimal?)Math.Max(0, -Delta.Value) : null; } }
        public string Location { get { return Sheet + " 第" + Row + "行"; } }
    }
    [DataContract] public class ExcelReviewState {
        [DataMember] public string FileName = "";
        [DataMember] public List<ReviewItem> Items = new List<ReviewItem>();
        [DataMember] public List<string> Messages = new List<string>();
        [DataMember] public decimal Threshold;
        [DataMember] public bool Detailed;
        [DataMember] public string LastText = "", LastTarget = "";
        [DataMember] public bool AutoLink;
        [DataMember] public int ModelVersion;
        [DataMember] public Dictionary<string, string> OriginalBodies;
        [DataMember] public bool CompactFeatures;
        [DataMember] public List<ReviewTotal> Totals;
        [DataMember] public List<ReviewBalance> Balances;
        [DataMember] public string WordingMode;
        [DataMember] public int AmountDecimals;
    }
    public class ExcelSheet {
        public string Name;
        public SortedDictionary<int, Dictionary<int, string>> Rows = new SortedDictionary<int, Dictionary<int, string>>();
        public List<int[]> Merges = new List<int[]>();
        public string Cell(int row, int col, bool merged = false) {
            Dictionary<int, string> cells; string text;
            if (Rows.TryGetValue(row, out cells) && cells.TryGetValue(col, out text)) return text;
            if (merged) foreach (var m in Merges) if (row >= m[0] && row <= m[2] && col >= m[1] && col <= m[3]) return Cell(m[0], m[1]);
            return "";
        }
    }
}
