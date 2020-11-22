using System;
using System.Collections.Generic;

namespace OrderManagementTool.Models
{
    public partial class LocationAddress
    {
        public long LocationAddressId { get; set; }
        public long LocationCodeId { get; set; }
        public string Address1 { get; set; }
        public string Address2 { get; set; }
        public string Address3 { get; set; }
        public string Address4 { get; set; }

        public virtual LocationCode LocationCode { get; set; }
    }
}
