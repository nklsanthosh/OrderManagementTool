using log4net;
using OrderManagementTool.Models;
using OrderManagementTool.Models.LogIn;
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

namespace OrderManagementTool
{
    /// <summary>
    /// Interaction logic for AddItemCategory.xaml
    /// </summary>
    public partial class AddItemCategory : Window
    {
        //ILog log = LogManager.GetLogger(typeof(MainWindow));
        OrderManagementContext orderManagementContext = new OrderManagementContext();
        private readonly Login _login;

        public AddItemCategory(Login login)
        {
            InitializeComponent();
            _login = login;
           // log.Info("In AddItemCategory Screen...");
            InitializeComponent();
            btn_add_category.IsEnabled = false;
            //txt_category_description.IsEnabled = false;
        }

        private void btn_add_category_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //log.Info("In add category...");
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
                //log.Info("Exception while adding category - " + ex.StackTrace);
                MessageBox.Show("An error occured during Item Category Creation. " + ex.Message, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
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
                    // txt_category_description.IsEnabled = false;
                     txt_category_description.Text = "";
                    MessageBox.Show("Item Category Name is already present", "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    txt_category_description.IsEnabled = true;
                    btn_add_category.IsEnabled = true;
                }
            }
        }
    }
}
