using System;
using System.Collections.Generic;
using System.Text;

namespace OrderManagementTool.Models.Purchase_Order
{
    public class ExportPO
    {
        public long POID;
        public long IndentID;
        public string Email;
        public string Remarks;
        public DateTime PODate;
        public List<Poitem> Poitems;
        public LocationAddress LocationAddressInfo;
        public string LocatioName;
    }
}
