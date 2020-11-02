using log4net;
using OrderManagementTool.Models;
using System.Windows;
using System.Linq;
using System;
using OrderManagementTool.Models.LogIn;

namespace OrderManagementTool
{
    /// <summary>
    /// Interaction logic for AddItemCode.xaml
    /// </summary>
    public partial class AddItemCode : Window
    {
        //ILog log = LogManager.GetLogger(typeof(MainWindow));
        OrderManagementContext orderManagementContext = new OrderManagementContext();
        private readonly Login _login;

        public AddItemCode(Login login)
        {
            _login = login;
            InitializeComponent();
            txt_item_code.IsEnabled = false;
            txt_item_description.IsEnabled = false;
            txt_item_technical_specification.IsEnabled = false;
            btn_add_item.IsEnabled = false;
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
                MessageBox.Show("Item Code " + itemCode + " is added successfully", "Order Management System", MessageBoxButton.OK, MessageBoxImage.Information);
                //log.Info("Item Code added...");
            }
            catch (Exception ex)
            {
               // log.Info("Exception while adding item code - " + ex.StackTrace);
                MessageBox.Show("An error occured during Item Code Creation. " + ex.Message, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
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
    }
}
