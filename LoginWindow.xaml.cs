using log4net;
using OrderManagementTool.Models;
using OrderManagementTool.Models.LogIn;
using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace OrderManagementTool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //ILog log = LogManager.GetLogger(typeof(MainWindow));
        public MainWindow()
        {
            //log.Info("In Login Screen...");
            InitializeComponent();
            btn_Login.IsEnabled = false;

        }
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //log.Info("Login Process started...");
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
                    //log.Error("Login error " + ex.StackTrace);
                }
                //log.Info("Login Process Completed...");
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

