﻿using System;
using System.Collections.Generic;

namespace OrderManagementTool.Models.Indent
{
    public class SaveIndent
    {
        public DateTime? Date { get; set; }
        public long LocationCode { get; set; }
        public long RaisedBy { get; set; }
        public string Email { get; set; }
        public DateTime CreateDate { get; set; }
        public long IndentId { get; set; }
        public string IndentRemarks { get; set; }
        public string ApprovalStatus { get; set; }
        public int RevisionNumber { get; set; }
        public long ApprovalID { get; set; }
        public string ApproverName { get; set; }

        public List<GridIndent> GridIndents { get; set; }
    }
}
