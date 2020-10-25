using System;
using System.Collections.Generic;
using System.Text;

namespace OrderManagementTool.Models.Indent
{
    public class ExcelIndent
    {
        // public int SlNo { get; set; }
        public string ItemCategoryName { get; set; }
        public string ItemCategoryDescription { get; set; }
        public string ItemName { get; set; }
        public string ItemCode { get; set; }
        public string Description { get; set; }
        //public int TotalPlanned { get; set; }
        //public int TotalSupplied { get; set; }
        public string Technical_Specifications { get; set; }
        public int Quantity { get; set; }
        //public int StockAsOn { get; set; }
        public string Units { get; set; }
        public string UnitsDescription { get; set; }
        public string Remarks { get; set; }
        public string CreatedBy { get; set; }
    }
}
