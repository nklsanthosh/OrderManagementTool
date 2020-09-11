using OrderManagementTool.Models;
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
                MessageBox.Show("Please Enter Valid UserName");
            }

            else if (txt_password.Password.Length == 0)
            {
                MessageBox.Show("Please Enter Valid Password");
            }

            else if (txt_username.Text.Length != 0 && txt_password.Password.Length != 0)
            {
                string username = txt_username.Text;
                string password = txt_password.Password.ToString();

                OrderManagementContext orderManagementContext = new OrderManagementContext();
                var userFound = orderManagementContext.UserMaster.Where(s => s.Email == username & s.Password == password).FirstOrDefault();

                if (userFound !=null)
                {
                    Menu menu = new Menu();
                    menu.Show();
                }
                else
                {
                    MessageBox.Show(" Enter Valid UserCredentials");
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
