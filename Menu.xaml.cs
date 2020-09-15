using OrderManagementTool.Models;
using OrderManagementTool.Models.LogIn;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Linq;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace OrderManagementTool
{
    /// <summary>
    /// Interaction logic for Menu.xaml
    /// </summary>
    public partial class Menu : Window
    {
        private readonly Login _login;
        OrderManagementContext orderManagementContext = new OrderManagementContext();
        public Menu()
        {
            InitializeComponent();
        }

        public Menu(Login login)
        {
            _login = login;
            InitializeComponent();
            try
            {
                var rolesMapping = (from roleMap in orderManagementContext.RoleMapping
                                    join roles in orderManagementContext.RolesMaster on roleMap.RoleId equals roles.RoleId
                                    join users in orderManagementContext.UserMaster on roleMap.UserId equals users.UserId
                                    where users.Email.Equals(_login.UserEmail)
                                    orderby roles.RoleId
                                    select new
                                    {
                                        roleMap.RoleId
                                    });
                if (rolesMapping != null)
                {
                    if (rolesMapping.FirstOrDefault().RoleId != 1)
                    {
                        lbl_item.Visibility = Visibility.Hidden;
                        btn_add_item.Visibility = Visibility.Hidden;
                    }
                    else
                    {
                        lbl_item.Visibility = Visibility.Visible;
                        btn_add_item.Visibility = Visibility.Visible;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured while fetching user information. "+ ex.Message, 
                                    "Order Management System", 
                                        MessageBoxButton.OK, 
                                            MessageBoxImage.Error);
            }
        }


        private void btn_create_Indent_Click(object sender, RoutedEventArgs e)
        {
            Indent indent = new Indent(_login);
            indent.Show();
            this.Close();
        }

        private void btn_add_item_Click(object sender, RoutedEventArgs e)
        {
            Item item = new Item(_login);
            item.Show();
            this.Close();
        }

        private void btn_log_off_Click(object sender, RoutedEventArgs e)
        {
            MainWindow mainWindow = new MainWindow();
            mainWindow.Show();
            this.Close();
        }
    }
}
