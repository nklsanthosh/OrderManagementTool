using System;
using System.Collections.Generic;
using System.Text;

namespace OrderManagementTool.Models.Indent
{
    public class SaveIndent
    {
        public DateTime? Date { get; set; }
        public string Location { get; set; }
        public long RaisedBy { get; set; }
        public DateTime CreateDate { get; set; }
        public long IndentId { get; set; }

        public long ApprovalID { get; set; }

        public List<GridIndent> GridIndents { get; set; }
    }
}
