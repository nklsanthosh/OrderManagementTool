using System;
using System.Collections.Generic;
using System.Text;

namespace OrderManagementTool.Models.Purchase_Order
{
    public class PORetrival
    {
        public long PO_ID { get; set; }
        public long Indent_No { get; set; }
        public long Approval_Status_Id { get; set; }
        public List<PoItemwithQuotation> poItemsWithQuotation;
      
    }
}
