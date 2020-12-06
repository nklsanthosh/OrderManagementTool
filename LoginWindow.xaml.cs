//using log4net;
using OrderManagementTool.Models;
using OrderManagementTool.Models.LogIn;
using System;
using System.Configuration;
using System.IO;
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
            ReadCredentials();
        }

        private void ReadCredentials()
        {
            try
            {
                string path = ConfigurationManager.AppSettings["InitializationPath"];

                DirectoryInfo dInfo = new DirectoryInfo(path);
                FileInfo[] files = null;
                try
                {
                    files = dInfo.GetFiles("OMT.ini");
                }
                catch (Exception e)
                {
                    Directory.CreateDirectory(path);
                    File.Create("OMT.ini");
                    files = dInfo.GetFiles("OMT.ini");
                }
                if (files.Length > 0)
                {
                    StreamReader file = File.OpenText(files[0].FullName);
                    string readInfo = file.ReadToEnd();
                    string userName = readInfo.Split(',')[0];
                    string password = readInfo.Split(',')[1];
                    txt_username.Text = userName;
                    txt_password.Password = password;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred while reading log in credentials :" + ex.Message,
                                      "Order Management System",
                                          MessageBoxButton.OK,
                                              MessageBoxImage.Error);
            }

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
                string path = ConfigurationManager.AppSettings["InitializationPath"];
                string username = txt_username.Text;
                string password = txt_password.Password.ToString().Split(new string[] { "\r\n" }, StringSplitOptions.None)[0];
                DirectoryInfo dInfo = new DirectoryInfo(path);
                FileInfo[] files = null;
                try
                {
                    files = dInfo.GetFiles("OMT.ini");
                }
                catch (Exception exc)
                {
                    string error = exc.Message;
                    Directory.CreateDirectory(path);
                    File.Create("OMT.ini");
                    files = dInfo.GetFiles("OMT.ini");
                }
                string[] lines = { username + "," + password };
                File.WriteAllLines(path + "\\OMT.ini", lines);

                try
                {
                    OrderManagementContext orderManagementContext = new OrderManagementContext();
                    //var userFound = orderManagementContext.UserMaster.Where(s => s.Email == username & s.Password == password).FirstOrDefault();
                    var userFound = (from i in orderManagementContext.UserMaster
                                     join emp in orderManagementContext.Employee
                                     on i.EmployeeId equals emp.EmployeeId
                                     where i.Email == username && i.Password == password
                                     select new
                                     {
                                         emp.FirstName,
                                         emp.LastName,
                                         emp.EmployeeId,
                                     }).FirstOrDefault();

                    if (userFound != null)
                    {
                        Login login = new Login();
                        login.UserEmail = username;
                        login.EmployeeID = userFound.EmployeeId;
                        login.UserName = userFound.FirstName + " " + userFound.LastName;

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

