using System;

namespace OrderManagementTool.Models
{
    public partial class Designation
    {
        public long DesignationId { get; set; }
        public long DepartmentId { get; set; }
        public string Designation1 { get; set; }
        public DateTime CreatedDate { get; set; }
        public long CreatedBy { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public long? ModifiedBy { get; set; }
    }
}
