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
    /// Interaction logic for AddUnit.xaml
    /// </summary>
    public partial class AddUnit : Window
    {
        ILog log = LogManager.GetLogger(typeof(MainWindow));
        OrderManagementContext orderManagementContext = new OrderManagementContext();
        private readonly Login _login;

        public AddUnit(Login login)
        {
            log.Info("In Add Unit Screen...");
            _login = login;
            InitializeComponent();
            btn_add_unit.IsEnabled = false;
            //txt_unit_description.IsEnabled = false;
        }

        private void txt_unit_LostFocus(object sender, RoutedEventArgs e)
        {
            log.Info("In Unit lost focus...");
            string unitname = txt_unit.Text;

            var isPresent = (from u in orderManagementContext.UnitMaster
                             where u.Unit == unitname
                             select u.UnitMasterId).FirstOrDefault();
            if (isPresent != 0)
            {
                txt_unit_description.Text = "";
                btn_add_unit.IsEnabled = false;
                MessageBox.Show("Unit is already present", "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                txt_unit_description.IsEnabled = true;
                btn_add_unit.IsEnabled = true;
            }
        }

        private void btn_add_unit_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                log.Info("Adding Units...");
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
                log.Info("Exception while adding units - " + ex.StackTrace);
                MessageBox.Show("An error occured during Unit Creation. " + ex.Message, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
