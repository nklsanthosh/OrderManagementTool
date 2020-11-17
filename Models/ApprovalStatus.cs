using System;
using System.Collections.Generic;

namespace OrderManagementTool.Models
{
    public partial class ApprovalStatus
    {
        public ApprovalStatus()
        {
            Poapproval = new HashSet<Poapproval>();
        }

        public long ApprovalStatusId { get; set; }
        public string ApprovalStatus1 { get; set; }
        public DateTime CreatedDate { get; set; }
        public long CreatedBy { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public long? ModifiedBy { get; set; }

        public virtual ICollection<Poapproval> Poapproval { get; set; }
    }
}
