using System;
using System.Collections.Generic;

namespace OrderManagementTool.Models
{
    public partial class IndentQuoteMapping
    {
        public long IndentQuoteId { get; set; }
        public int IndentId { get; set; }
        public int QuotationId { get; set; }
        public int CreatedBy { get; set; }
        public DateTime CreatedDate { get; set; }
        public int? ModifiedBy { get; set; }
        public DateTime? ModifiedDate { get; set; }
    }
}
