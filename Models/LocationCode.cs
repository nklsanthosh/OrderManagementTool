﻿using System;
using System.Collections.Generic;

namespace OrderManagementTool.Models
{
    public partial class LocationCode
    {
        public LocationCode()
        {
            LocationAddress = new HashSet<LocationAddress>();
        }

        public long LocationCodeId { get; set; }
        public long LocationId { get; set; }
        public string LocationName { get; set; }

        public virtual ICollection<LocationAddress> LocationAddress { get; set; }
    }
}
