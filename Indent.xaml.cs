using OrderManagementTool.Models;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Linq;
using Caliburn.Micro;

namespace OrderManagementTool
{
    /// <summary>
    /// Interaction logic for Indent.xaml
    /// </summary>
    public partial class Indent : Window
    {

        // private readonly List<string> itemName;
        OrderManagementContext orderManagementContext = new OrderManagementContext();

        public BindableCollection<string> ItemName { get; set; }
        public Indent()
        {
            InitializeComponent();
        }

        private void DataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void cbx_itemname_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var itemName = (from i in orderManagementContext.ItemCategory
                            select i.ItemCategoryName).Distinct().ToList();
            foreach (var i in itemName)
            {
                cbx_itemname.Items.Add(i.Trim());
            }
        }

        private void cbx_itemname_DropDownOpened(object sender, EventArgs e)
        {
            cbx_itemname.Items.Clear();
            var itemName = (from i in orderManagementContext.ItemCategory
                            select i.ItemCategoryName).Distinct().ToList();
            foreach (var i in itemName)
            {
                cbx_itemname.Items.Add(i.Trim());
            }

            //cbx_itemname.DisplayMemberPath = "Text";
        }

        private void cbx_itemcode_DropDownOpened(object sender, EventArgs e)
        {
            var itemName = cbx_itemname.SelectedItem.ToString();

            if (itemName != null)
            {
                cbx_itemcode.Items.Clear();
                var itemCode = (from i in orderManagementContext.ItemMaster
                                from ic in orderManagementContext.ItemCategory
                                where i.ItemCategoryId == ic.ItemCategoryId && ic.ItemCategoryName == itemName
                                select i.ItemName).Distinct().ToList();
                foreach (var i in itemCode)
                {
                    cbx_itemcode.Items.Add(i);
                }
            }
            else
            {
                MessageBox.Show("Please Select ItemName");
            }
        }

        private void txt_quantity_changed(object sender, SelectionChangedEventArgs e)
        {
            if (Int32.TryParse(txt_quantity.Text, out int value))
            {
                var i = 0;
                // Here comes the code if numeric
            }
            else
            {
                MessageBox.Show("Please enter Valid Number");
            }
            // var quantity = txt_quantity.Text;
        }
        private void cbx_itemcode_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
        }

        private void cbx_units_DropDownOpened(object sender, EventArgs e)
        {
            cbx_units.Items.Clear();

            var units = (from u in orderManagementContext.UnitMaster
                         select u.Unit).Distinct().ToList();
            foreach (var u in units)
            {
                cbx_units.Items.Add(u.Trim());
            }
        }
    }
}
