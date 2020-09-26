using System;

namespace OrderManagementTool.Models
{
    public partial class IndentApproval
    {
        public long IndentApprovalId { get; set; }
        public long IndentId { get; set; }
        public long ApprovalId { get; set; }
        public long ApprovalStatusId { get; set; }
        public string Remarks { get; set; }
        public DateTime CreatedDate { get; set; }
        public long CreatedBy { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public long? ModifiedBy { get; set; }
    }
}
