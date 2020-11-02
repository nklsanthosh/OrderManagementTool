//using log4net;
using OrderManagementTool.Models;
using OrderManagementTool.Models.LogIn;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace OrderManagementTool
{

    /// <summary>
    /// Interaction logic for Item.xaml
    /// </summary>
    public partial class Item : Window
    {
       //ILog log = LogManager.GetLogger(typeof(MainWindow));
        OrderManagementContext orderManagementContext = new OrderManagementContext();
        private readonly Login _login;

        public Item(Login login)
        {
            _login = login;
            //log.Info("In Masters Screen...");
            InitializeComponent();
            txt_item_code.IsEnabled = false;
            txt_item_description.IsEnabled = false;
            txt_item_technical_specification.IsEnabled = false;
            btn_add_item.IsEnabled = false;
            btn_add_unit.IsEnabled = false;
            txt_unit_description.IsEnabled = false;
            btn_add_category.IsEnabled = false;
            txt_category_description.IsEnabled = false;
            txt_item_name.IsEnabled = false;
        }

        private void txt_label_name_LostFocus(object sender, RoutedEventArgs e)
        {
            //log.Info("In Item Name lost focus...");
            string itemName = txt_item_name.Text;
            if (itemName != "")
            {
                var isPresent = (from i in orderManagementContext.ItemMaster
                                 where i.ItemName == itemName
                                 select i.ItemCategoryId).FirstOrDefault();

                if (isPresent != 0)
                {
                    txt_item_code.Text = "";
                    txt_item_code.IsEnabled = false;
                    txt_item_description.IsEnabled = false;
                    txt_item_technical_specification.IsEnabled = false;
                    btn_add_item.IsEnabled = false;
                    MessageBox.Show("Item Name is already present", "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    txt_item_code.IsEnabled = true;
                }
            }
        }

        private void txt_item_code_LostFocus(object sender, RoutedEventArgs e)
        {
            //log.Info("In item code lost focus...");
            string labelName = txt_item_name.Text;
            string labelCode = txt_item_code.Text;

            if (labelCode != "")
            {
                var isPresent = (from i in orderManagementContext.ItemMaster
                                 from ic in orderManagementContext.ItemCategory
                                 where i.ItemCategoryId == ic.ItemCategoryId && ic.ItemCategoryName == labelName && i.ItemCode == labelCode
                                 select i.ItemMasterId).FirstOrDefault();

                if (isPresent != 0)
                {
                    txt_item_description.IsEnabled = false;
                    txt_item_technical_specification.IsEnabled = false;
                    btn_add_item.IsEnabled = false;
                    MessageBox.Show("Item Code is already present", "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    txt_item_code.IsEnabled = true;
                    btn_add_item.IsEnabled = true;
                    txt_item_description.IsEnabled = true;
                    txt_item_technical_specification.IsEnabled = true;
                }
            }
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void txt_unit_LostFocus(object sender, RoutedEventArgs e)
        {
            //log.Info("In Unit lost focus...");
            string unitname = txt_unit.Text;

            var isPresent = (from u in orderManagementContext.UnitMaster
                             where u.Unit == unitname
                             select u.UnitMasterId).FirstOrDefault();
            if (isPresent != 0)
            {
                txt_unit_description.IsEnabled = false;
                btn_add_unit.IsEnabled = false;
                MessageBox.Show("Unit is already present", "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                txt_unit_description.IsEnabled = true;
                btn_add_unit.IsEnabled = true;
            }
        }

        private void txt_category_name_LostFocus(object sender, RoutedEventArgs e)
        {
            //log.Info("In item Category lost focus...");
            string categoryName = txt_category_name.Text;
            if (categoryName != "")
            {
                var isPresent = (from i in orderManagementContext.ItemCategory
                                 where i.ItemCategoryName == categoryName
                                 select i.ItemCategoryId).FirstOrDefault();
                if (isPresent != 0)
                {
                    btn_add_category.IsEnabled = false;
                    txt_category_description.IsEnabled = false;
                    MessageBox.Show("Item Category Name is already present", "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    txt_category_description.IsEnabled = true;
                    btn_add_category.IsEnabled = true;
                }
            }
        }

        private void btn_add_category_Click(object sender, RoutedEventArgs e)
        {
            try
            {
               // log.Info("In add category...");
                string categoryName = txt_category_name.Text;
                string categoryDescription = txt_category_description.Text;

                orderManagementContext.ItemCategory.Add(
                    new ItemCategory
                    {
                        ItemCategoryName = categoryName,
                        CreatedDate = DateTime.Now,
                        CreatedBy = _login.EmployeeID,
                        Description = categoryDescription
                    }
                    );
                orderManagementContext.SaveChanges();
                txt_category_name.Text = "";
                txt_category_description.Text = "";
                MessageBox.Show("Item Category " + categoryName + " is added successfully", "Order Management System", MessageBoxButton.OK, MessageBoxImage.Information);
                //log.Info("Category added...");
            }
            catch (Exception ex)
            {
                //log.Info("Exception while adding category - " +ex.StackTrace);
                MessageBox.Show("An error occured during Item Category Creation. " + ex.Message, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void cbx_category_name_DropDownOpened(object sender, EventArgs e)
        {
            cbx_category_name.Items.Clear();
            var categoryName = (from i in orderManagementContext.ItemCategory
                                select i.ItemCategoryName).Distinct().ToList();
            foreach (var i in categoryName)
            {
                cbx_category_name.Items.Add(i.Trim());
            }
        }

        private void btn_add_item_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //log.Info("Adding Item Code...");
                string itemName = txt_item_name.Text;
                string itemCode = txt_item_code.Text;
                string itemDescription = txt_item_description.Text;
                string itemtechnicalSpec = txt_item_technical_specification.Text;
                string categoryName = cbx_category_name.SelectedItem.ToString();
                var itemCategoryId = (from i in orderManagementContext.ItemCategory
                                      where i.ItemCategoryName == categoryName
                                      select i.ItemCategoryId).FirstOrDefault();

                orderManagementContext.ItemMaster.Add(
                    new ItemMaster
                    {
                        ItemCode = itemCode,
                        ItemName = itemName,
                        Description = itemDescription,
                        TechnicalSpecification = itemtechnicalSpec,
                        ItemCategoryId = itemCategoryId,
                        CreatedDate = DateTime.Now,
                        CreatedBy = _login.EmployeeID
                    });
                orderManagementContext.SaveChanges();
                MessageBox.Show("Item Code " + itemCode + " is added successfully", "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                //log.Info("Item Code added...");
            }
            catch (Exception ex)
            {
                //log.Info("Exception while adding item code - " + ex.StackTrace);
                MessageBox.Show("An error occured during Item Code Creation. " + ex.Message, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void cbx_category_name_DropDownClosed(object sender, EventArgs e)
        {
            if (cbx_category_name.SelectedValue == null)
            {
                txt_item_name.IsEnabled = false;
                txt_item_name.Text = "";
                txt_item_code.Text = "";
                txt_item_code.IsEnabled = false;
                txt_item_description.IsEnabled = false;
                txt_item_description.Text = "";
                txt_item_technical_specification.IsEnabled = false;
                txt_item_technical_specification.Text = "";
                btn_add_item.IsEnabled = false;
            }
            else
            {
                txt_item_name.IsEnabled = true;
            }
        }

        private void btn_add_unit_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //log.Info("Adding Units...");
                string unit = txt_unit.Text;
                string unitDescription = txt_unit_description.Text;
                orderManagementContext.UnitMaster.Add(
                    new UnitMaster
                    {
                        Unit = unit,
                        Description = unitDescription,
                        CreatedBy = _login.EmployeeID,
                        CreatedDate = DateTime.Now
                    });
                orderManagementContext.SaveChanges();
                MessageBox.Show("Unit " + unit + " is added successfully", "Order Management System", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                //log.Info("Exception while adding units - " + ex.StackTrace);
                MessageBox.Show("An error occured during Unit Creation. " + ex.Message, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            Menu menu = new Menu(_login);
            menu.Show();
        }
    }
}

