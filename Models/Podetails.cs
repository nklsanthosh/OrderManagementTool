using System;
using System.Collections.Generic;

namespace OrderManagementTool.Models
{
    public partial class Podetails
    {
        public long PodetailsId { get; set; }
        public long PoId { get; set; }
        public string OfferNumber { get; set; }
        public string VendorName { get; set; }
        public string VendorCode { get; set; }
        public DateTime OfferDate { get; set; }
        public string ContactNumber { get; set; }
        public string ContactPerson { get; set; }
        public string Description { get; set; }
        public long Quantity { get; set; }
        public long Units { get; set; }
        public decimal UnitPrice { get; set; }
        public decimal GstValue { get; set; }
        public decimal TotalPrice { get; set; }
        public int QNo { get; set; }
        public DateTime CreatedDate { get; set; }
        public long CreatedBy { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public long? ModifiedBy { get; set; }

        public virtual UnitMaster UnitsNavigation { get; set; }
    }
}
