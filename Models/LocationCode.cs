using System;
using System.Collections.Generic;

namespace OrderManagementTool.Models
{
    public partial class LocationCode
    {
        public long LocationCodeId { get; set; }
        public long LocationId { get; set; }
        public string LocationName { get; set; }
    }
}
