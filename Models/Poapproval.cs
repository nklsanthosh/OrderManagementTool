using System;
using System.Collections.Generic;

namespace OrderManagementTool.Models
{
    public partial class Poapproval
    {
        public long PoapprovalId { get; set; }
        public long PoId { get; set; }
        public long ApprovalId { get; set; }
        public long ApprovalStatusId { get; set; }
        public DateTime CreatedDate { get; set; }
        public long? CreatedBy { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public long? ModifiedBy { get; set; }

        public virtual ApprovalStatus ApprovalStatus { get; set; }
    }
}
