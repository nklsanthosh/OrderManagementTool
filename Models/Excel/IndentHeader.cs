using System;

namespace OrderManagementTool.Models.Excel
{
    public class IndentHeader
    {
        public string To { get; set; }
        public string AddressLine1 { get; set; }
        public string AddressLine2 { get; set; }
        public string AddressLine3 { get; set; }
        public string Contact { get; set; }
        public string IndentNo { get; set; }
        public DateTime IndentDate { get; set; }
        public string RefNo { get; set; }
        public DateTime RefDate { get; set; }
        public string Remarks { get; set; }
        public string Attention { get; set; }
        public string GSTNo { get; set; }
        public string IECNo { get; set; }
        public string Project { get; set; }
        public string WBS { get; set; }
    }
}
