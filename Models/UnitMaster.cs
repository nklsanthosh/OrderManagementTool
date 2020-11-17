using System;
using System.Collections.Generic;

namespace OrderManagementTool.Models
{
    public partial class UnitMaster
    {
        public UnitMaster()
        {
            Podetails = new HashSet<Podetails>();
        }

        public long UnitMasterId { get; set; }
        public string Unit { get; set; }
        public string Description { get; set; }
        public DateTime CreatedDate { get; set; }
        public long CreatedBy { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public long? ModifiedBy { get; set; }

        public virtual ICollection<Podetails> Podetails { get; set; }
    }
}
