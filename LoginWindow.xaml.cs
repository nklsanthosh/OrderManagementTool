using OrderManagementTool.Models;
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
                try
                {
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
                        var splashScreen = new SplashScreen();
                        splashScreen.Close();
                        this.Close();
                    }
                    else
                    {
                        MessageBox.Show("Enter Valid UserCredentials", "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occurred while logging in :" + ex.Message,
                                       "Order Management System",
                                           MessageBoxButton.OK,
                                               MessageBoxImage.Error);
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
        //private void txt_password_TextChanged(object sender, TextChangedEventArgs e)
        //{
        //    if (txt_password.Password.Length == 0)
        //    {
        //        txt_password.Template.Triggers;
        //    }
        //    else
        //    {
        //        btn_Login.IsEnabled = true;
        //    }
        //}

    }
    public class PasswordBoxMonitor : DependencyObject
    {
        public static bool GetIsMonitoring(DependencyObject obj)
        {
            return (bool)obj.GetValue(IsMonitoringProperty);
        }

        public static void SetIsMonitoring(DependencyObject obj, bool value)
        {
            obj.SetValue(IsMonitoringProperty, value);
        }

        public static readonly DependencyProperty IsMonitoringProperty =
            DependencyProperty.RegisterAttached("IsMonitoring", typeof(bool), typeof(PasswordBoxMonitor), new UIPropertyMetadata(false, OnIsMonitoringChanged));



        public static int GetPasswordLength(DependencyObject obj)
        {
            return (int)obj.GetValue(PasswordLengthProperty);
        }

        public static void SetPasswordLength(DependencyObject obj, int value)
        {
            obj.SetValue(PasswordLengthProperty, value);
        }

        public static readonly DependencyProperty PasswordLengthProperty =
            DependencyProperty.RegisterAttached("PasswordLength", typeof(int), typeof(PasswordBoxMonitor), new UIPropertyMetadata(0));

        private static void OnIsMonitoringChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var pb = d as PasswordBox;
            if (pb == null)
            {
                return;
            }
            if ((bool)e.NewValue)
            {
                pb.PasswordChanged += PasswordChanged;
            }
            else
            {
                pb.PasswordChanged -= PasswordChanged;
            }
        }

        static void PasswordChanged(object sender, RoutedEventArgs e)
        {
            var pb = sender as PasswordBox;
            if (pb == null)
            {
                return;
            }
            SetPasswordLength(pb, pb.Password.Length);
        }
    }
}

