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

        private List<string> itemName;
        private List<string> itemCode;
        private string selectedItemCode;
        private List<string> units;
        private string selectedItemName;
        private int quantityEntered;
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
            itemName = (from i in orderManagementContext.ItemCategory
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
        }

        private void cbx_itemcode_DropDownOpened(object sender, EventArgs e)
        {
            if (cbx_itemname.SelectedItem != null)
            {
                selectedItemName = cbx_itemname.SelectedItem.ToString();
                cbx_itemcode.Items.Clear();
                itemCode = (from i in orderManagementContext.ItemMaster
                            from ic in orderManagementContext.ItemCategory
                            where i.ItemCategoryId == ic.ItemCategoryId && ic.ItemCategoryName == selectedItemName
                            select i.ItemCode).Distinct().ToList();
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

        private void cbx_units_DropDownOpened(object sender, EventArgs e)
        {
            cbx_units.Items.Clear();
            units = (from u in orderManagementContext.UnitMaster
                     select u.Unit).Distinct().ToList();
            foreach (var u in units)
            {
                cbx_units.Items.Add(u.Trim());
            }
        }

        private void txt_quantity_lostfocus(object sender, RoutedEventArgs e)
        {
            if (Int32.TryParse(txt_quantity.Text, out int value))
            {
                quantityEntered = Convert.ToInt32(txt_quantity.Text);
            }
            else
            {
                MessageBox.Show("Please enter Valid Number");
            }
        }

        private void cbx_itemcode_DropDownClosed(object sender, EventArgs e)
        {
            selectedItemCode = cbx_itemcode.SelectionBoxItem.ToString();
            var itemDetails = (from i in orderManagementContext.ItemMaster
                               from ic in orderManagementContext.ItemCategory
                               where i.ItemCategoryId == ic.ItemCategoryId && ic.ItemCategoryName == selectedItemName && i.ItemCode == selectedItemCode
                               select i).FirstOrDefault();
            if (itemDetails != null)
            {
                txt_description.Text = itemDetails.Description;
                txt_technical_description.Text = itemDetails.TechnicalSpecification;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (cbx_itemname.SelectedItem != null && cbx_itemcode.SelectedItem != null && txt_quantity.Text != "" 
               && quantityEntered != 0 && cbx_units.SelectedItem != null)
            {

            }
            else if (cbx_itemname.SelectedItem == null)
            {
                MessageBox.Show("Please Select ItemName");
            }
            else if (cbx_itemcode.SelectedItem == null)
            {
                MessageBox.Show("Please Select ItemCode");
            }
            else if (txt_quantity.Text == "")
            {
                MessageBox.Show("Please Enter Quantity");
            }
            else if (quantityEntered == 0)
            {
                MessageBox.Show("Please Enter Valid Quantity");
            }
            else if (cbx_units.SelectedItem == null)
            {
                MessageBox.Show("Please Select Units");
            }
        }
    }
}
