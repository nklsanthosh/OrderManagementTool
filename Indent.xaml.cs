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
using OrderManagementTool.Models.Indent;
using System.Data;

namespace OrderManagementTool
{
    /// <summary>
    /// Interaction logic for Indent.xaml
    /// </summary>
    public partial class Indent : Window
    {
        private List<string> itemName;
        private List<string> itemCode;
        private string selectedItemCode;
        private List<string> units;
        private string selectedItemName;
        private int quantityEntered;
        private int gridSelectedIndex;
        OrderManagementContext orderManagementContext = new OrderManagementContext();
        private List<GridIndent> gridIndents = new List<GridIndent>();
        public BindableCollection<string> ItemName { get; set; }
        public Indent()
        {
            InitializeComponent();
            LoadItemName();
        }

        private void DataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //DataRowView rowview = grid_indentdata.SelectedItem as DataRowView;
            var rowview = grid_indentdata.SelectedItem as GridIndent;
            gridSelectedIndex = grid_indentdata.SelectedIndex;

            if (rowview != null)
            {
                txt_quantity.Text = Convert.ToInt32(rowview.Quantity).ToString();
                LoadItemName();
                cbx_itemname.SelectedItem = rowview.ItemName;
                LoadItemCode();
                cbx_itemcode.SelectedItem = rowview.ItemCode;
                LoadDesctiption();
                LoadUnits();
                cbx_units.SelectedItem = rowview.Units;
            }
        }

        private void cbx_itemname_DropDownOpened(object sender, EventArgs e)
        {
            LoadItemName();
        }

        private void cbx_itemcode_DropDownOpened(object sender, EventArgs e)
        {
            LoadItemCode();
        }

        private void cbx_units_DropDownOpened(object sender, EventArgs e)
        {
            LoadUnits();
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
            LoadDesctiption();

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (cbx_itemname.SelectedItem != null && cbx_itemcode.SelectedItem != null && txt_quantity.Text != ""
               && quantityEntered != 0 && cbx_units.SelectedItem != null)
            {
                bool itemPresent = false;
                GridIndent gridIndent = new GridIndent();
                gridIndent.SlNo = gridIndents.Count + 1;
                gridIndent.ItemCode = cbx_itemcode.SelectedItem.ToString();
                gridIndent.ItemName = cbx_itemname.SelectedItem.ToString();
                gridIndent.Quantity = quantityEntered;
                gridIndent.Technical_Specifications = txt_technical_description.Text.Trim();
                gridIndent.Units = cbx_units.SelectedItem.ToString();
                gridIndent.Remarks = txt_remarks.Text.Trim();

                foreach (var i in gridIndents)
                {
                    if (i.ItemName == gridIndent.ItemName && i.ItemCode == gridIndent.ItemCode && i.Description == gridIndent.Description && i.Technical_Specifications == gridIndent.Technical_Specifications && itemPresent == false)
                    {
                        itemPresent = true;

                    }
                }
                if (!itemPresent)
                {
                    gridIndents.Add(gridIndent);
                    grid_indentdata.ItemsSource = null;
                    grid_indentdata.ItemsSource = gridIndents;
                    ClearFields();
                }
                else
                {
                    MessageBox.Show("Item Already Present");
                }
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

        private void LoadItemName()
        {
            cbx_itemname.Items.Clear();
            var itemName = (from i in orderManagementContext.ItemCategory
                            select i.ItemCategoryName).Distinct().ToList();
            foreach (var i in itemName)
            {
                cbx_itemname.Items.Add(i.Trim());
            }
        }
        private void LoadDesctiption()
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
        private void LoadUnits()
        {
            cbx_units.Items.Clear();
            units = (from u in orderManagementContext.UnitMaster
                     select u.Unit).Distinct().ToList();
            foreach (var u in units)
            {
                cbx_units.Items.Add(u.Trim());
            }
        }
        private void LoadItemCode()
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
        private void ClearFields()
        {
            cbx_itemname.SelectedValue = "";
            cbx_itemcode.SelectedValue = "";
            txt_description.Text = "";
            txt_quantity.Text = "";
            txt_technical_description.Text = "";
            cbx_units.SelectedValue = "";
            txt_remarks.Text = "";
            txt_technical_description.Text = "";
        }

        private void btn_clear_fields_Click(object sender, RoutedEventArgs e)
        {
            ClearFields();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (gridSelectedIndex >= 0)
            {
                bool itemPresent = false;
                GridIndent value = new GridIndent();
                value.ItemCode = cbx_itemcode.SelectedItem.ToString();
                value.ItemName = cbx_itemname.SelectedItem.ToString();
                value.Quantity = quantityEntered;
                value.Technical_Specifications = txt_technical_description.Text.Trim();
                value.Units = cbx_units.SelectedItem.ToString();
                value.Remarks = txt_remarks.Text.Trim();

                foreach (var i in gridIndents)
                {
                    if (i.ItemName == value.ItemName && i.ItemCode == value.ItemCode && i.Description == value.Description && i.Technical_Specifications == value.Technical_Specifications && itemPresent == false)
                    {
                        itemPresent = true;
                    }
                }
                if (!itemPresent)
                {
                    var gridIndent = gridIndents[gridSelectedIndex];
                    gridIndent.ItemCode = cbx_itemcode.SelectedItem.ToString();
                    gridIndent.ItemName = cbx_itemname.SelectedItem.ToString();
                    gridIndent.Quantity = quantityEntered;
                    gridIndent.Technical_Specifications = txt_technical_description.Text.Trim();
                    gridIndent.Units = cbx_units.SelectedItem.ToString();
                    gridIndent.Remarks = txt_remarks.Text.Trim();
                    grid_indentdata.ItemsSource = null;
                    grid_indentdata.ItemsSource = gridIndents;
                    gridSelectedIndex = -1;
                    ClearFields();
                }
                else
                {
                    MessageBox.Show("Item Already Present");
                }
            }
        }
    }
}
