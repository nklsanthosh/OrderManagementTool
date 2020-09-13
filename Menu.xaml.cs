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

namespace OrderManagementTool
{
    /// <summary>
    /// Interaction logic for Menu.xaml
    /// </summary>
    public partial class Menu : Window
    {
        private readonly Login _login;
        public Menu()
        {
            InitializeComponent();
        }

        public Menu(Login login)
        {
            _login = login;
            InitializeComponent();
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
