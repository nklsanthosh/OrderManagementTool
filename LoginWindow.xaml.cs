﻿using OrderManagementTool.Models;
using OrderManagementTool.Models.LogIn;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace OrderManagementTool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            btn_Login.IsEnabled = false;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {

            if (txt_username.Text.Length == 0)
            {
                MessageBox.Show("Please Enter Valid UserName", "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            else if (txt_password.Password.Length == 0)
            {
                MessageBox.Show("Please Enter Valid Password", "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            else if (txt_username.Text.Length != 0 && txt_password.Password.Length != 0)
            {
                string username = txt_username.Text;
                string password = txt_password.Password.ToString();

                OrderManagementContext orderManagementContext = new OrderManagementContext();
                //var userFound = orderManagementContext.UserMaster.Where(s => s.Email == username & s.Password == password).FirstOrDefault();
                var userFound = (from i in orderManagementContext.UserMaster
                                 where i.Email == username && i.Password == password
                                 select i).FirstOrDefault();

                if (userFound != null)
                {
                    Login login = new Login();
                    login.UserEmail = username;
                    login.EmployeeID = userFound.EmployeeId;
                    Menu menu = new Menu(login);
                    menu.Show();
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Enter Valid UserCredentials", "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void txt_username_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (txt_username.Text.Length == 0 && txt_password.Password.Length == 0)
            {
                btn_Login.IsEnabled = false;
            }
            else
            {
                btn_Login.IsEnabled = true;
            }
        }
    }
}
