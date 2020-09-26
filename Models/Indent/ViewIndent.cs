using System;

namespace OrderManagementTool.Models.Indent
{
    public class ViewIndent
    {
        public int Sl_No { get; set; }
        public long IndentId { get; set; }
        public string Email { get; set; }
        public DateTime? Date { get; set; }
        public string Location { get; set; }
        public string ApproverName { get; set; }
        public string IndentRemarks { get; set; }
        public string CategoryName { get; set; }
        public string ItemCode { get; set; }
        public string Description { get; set; }
        public int TotalPlanned { get; set; }
        public int TotalSupplied { get; set; }
        public string Technical_Specifications { get; set; }
        public int Quantity { get; set; }
        public int StockAsOn { get; set; }
        public string Units { get; set; }
        public string Remarks { get; set; }

    }
}
