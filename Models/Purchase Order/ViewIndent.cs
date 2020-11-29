using System;
using System.Collections.Generic;
using System.Text;

namespace OrderManagementTool.Models.Purchase_Order
{
    public class ViewIndent
    {
        public int Sl_No { get; set; }
        public string Email { get; set; }
        public DateTime? Date { get; set; }
        public string Location { get; set; }
        public long LocationId { get; set; }
        public string ApproverName { get; set; }
        public string IndentRemarks { get; set; }
        public string CategoryName { get; set; }
        public string ItemCode { get; set; }
        public string Description { get; set; }
        public string Technical_Specifications { get; set; }
        public int Quantity { get; set; }
        public string Units { get; set; }
        public string Remarks { get; set; }
    }
}
