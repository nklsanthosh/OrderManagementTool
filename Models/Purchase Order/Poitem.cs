using System;
using System.Collections.Generic;
using System.Text;

namespace OrderManagementTool.Models.Purchase_Order
{
    public class Poitem
    {
        public int Sl_NO { get; set; }
        public string Description { get; set; }
        public int Quantity { get; set; }
        public string Units { get; set; }
        public double Unit_Price { get; set; }
        public double Total_Price { get; set; }
    }
}
