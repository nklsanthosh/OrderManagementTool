using System;
using System.Collections.Generic;

namespace OrderManagementTool.Models
{
    public partial class QuotationInformation
    {
        public long QuotationId { get; set; }
        public int? QuoteNo { get; set; }
        public string QuotationFileName { get; set; }
        public string FileContents { get; set; }
        public long CreatedBy { get; set; }
        public DateTime CreatedDate { get; set; }
        public long? ModifiedBy { get; set; }
        public DateTime? ModifedDate { get; set; }
    }
}
