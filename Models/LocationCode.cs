using System;
using System.Collections.Generic;

namespace OrderManagementTool.Models
{
    public partial class LocationCode
    {
        public LocationCode()
        {
            IndentMaster = new HashSet<IndentMaster>();
        }

        public long LocationId { get; set; }
        public string LocationName { get; set; }

        public virtual ICollection<IndentMaster> IndentMaster { get; set; }
    }
}
