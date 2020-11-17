using System;
using System.Collections.Generic;

namespace OrderManagementTool.Models
{
    public partial class Pomaster
    {
        public long PoId { get; set; }
        public long IndentId { get; set; }
        public string Remarks { get; set; }
        public DateTime CreatedDate { get; set; }
        public long CreatedBy { get; set; }
        public DateTime? ModifiedDate { get; set; }
        public long? ModifiedBy { get; set; }
    }
}
