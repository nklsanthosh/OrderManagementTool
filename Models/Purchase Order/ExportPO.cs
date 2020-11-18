using System;
using System.Collections.Generic;
using System.Text;

namespace OrderManagementTool.Models.Purchase_Order
{
    public class ExportPO
    {
        public int POID;
        public int IndentID;
        public string Email;
        public DateTime PODate;
        public List<Poitem> Poitems;
    }
}
