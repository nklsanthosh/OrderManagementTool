﻿using System;
using System.Collections.Generic;
using System.Text;

namespace OrderManagementTool.Models.Purchase_Order
{
    public class POView
    {
        public int Sl_No { get; set; }
        public long PO_ID { get; set; }
        public long Indent_No { get; set; }
        public long Approval_Status_Id { get; set; }
        public long Approver_Id { get; set; }
        public string Offer_Number { get; set; }
        public string Vendor_Name { get; set; }
        public string Vendor_Code { get; set; }
        public string Offer_Date { get; set; }
        public string Contact_No { get; set; }
        public string Contact_Person { get; set; }
        public string Description { get; set; }
        public int Quantity { get; set; }
        public string Units { get; set; }
        public double Unit_Price { get; set; }
        public decimal GST_Value { get; set; }
        public decimal Total_Price { get; set; }
        public int Q_No { get; set; }
    }
}
