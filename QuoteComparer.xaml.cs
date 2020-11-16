using ExcelDataReader;
using Microsoft.Data.SqlClient;
using Microsoft.Win32;
using OrderManagementTool.Models;
using OrderManagementTool.Models.Indent;
using OrderManagementTool.Models.LogIn;
using OrderManagementTool.Models.Purchase_Order;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.IO;
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
    /// Interaction logic for QuoteComparer.xaml
    /// </summary>
    public partial class QuoteComparer : Window
    {
        //ILog log = LogManager.GetLogger(typeof(MainWindow));
        private string path = Convert.ToString(ConfigurationManager.AppSettings["InputFilePath"]);
        //private string targetPath = Convert.ToString(ConfigurationManager.AppSettings["TargetReportPath"]);
        private static List<ViewIndent> viewIndents = new List<ViewIndent>();
        private List<string> itemCode;
        private string selectedItemCode;
        private List<string> units;
        private string selectedItemCategoryName;
        private int quantityEntered;
        private int gridSelectedIndex;
        OrderManagementContext orderManagementContext = new OrderManagementContext();
        private List<GridIndent> gridIndents = new List<GridIndent>();
        // public BindableCollection<string> ItemName { get; set; }
        private readonly Login _login;
        private static string filePathLocation;
        public long indentNo;
        // public string mailFrom;
        public string mailTo;
        private bool isGridReadOnly = false;
        List<Poitem> poitems_1 = new List<Poitem>();
        List<Poitem> poitems_2 = new List<Poitem>();
        List<Poitem> poitems_3 = new List<Poitem>();

        public QuoteComparer(Login login)
        {
            _login = login;
            InitializeComponent();
        }

        private void grid_po_confirmation_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void btn_upload_Click_1(object sender, RoutedEventArgs e)
        {
            try
            {
                string filePath = string.Empty;
                OpenFileDialog openFileDialog = new OpenFileDialog();
                if (openFileDialog.ShowDialog() == true)
                    filePath = openFileDialog.FileName;

                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    poitems_1 = PullIndentData(stream);
                }
                grid_quote_1.ItemsSource = null;
                grid_quote_1.ItemsSource = poitems_1;
                File.SetAttributes(filePath, FileAttributes.Normal);
                string targetPath = Convert.ToString(ConfigurationManager.AppSettings["TargetReportPath"]);

                targetPath = targetPath + "\\" + openFileDialog.SafeFileName;
                File.Move(filePath, targetPath, true);
                MessageBox.Show("The file has been processed and data has been uploaded.",
                        "Order Management System", MessageBoxButton.OK, MessageBoxImage.Information);
                // log.Info("The file has been processed and data has been uploaded.");
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to upload the data from the file. An error occured : " + ex.Message,
                    "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                // log.Error("Unable to upload the data from the file. An error occured : " + ex.Message);
            }
        }

        private void btn_upload_Click_2(object sender, RoutedEventArgs e)
        {
            try
            {
                string filePath = string.Empty;
                OpenFileDialog openFileDialog = new OpenFileDialog();
                if (openFileDialog.ShowDialog() == true)
                    filePath = openFileDialog.FileName;

                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    poitems_2 = PullIndentData(stream);
                }
                grid_quote_2.ItemsSource = null;
                grid_quote_2.ItemsSource = poitems_2;
                File.SetAttributes(filePath, FileAttributes.Normal);
                string targetPath = Convert.ToString(ConfigurationManager.AppSettings["TargetReportPath"]);

                targetPath = targetPath + "\\" + openFileDialog.SafeFileName;
                File.Move(filePath, targetPath, true);
                MessageBox.Show("The file has been processed and data has been uploaded.",
                        "Order Management System", MessageBoxButton.OK, MessageBoxImage.Information);
                // log.Info("The file has been processed and data has been uploaded.");
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to upload the data from the file. An error occured : " + ex.Message,
                    "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                // log.Error("Unable to upload the data from the file. An error occured : " + ex.Message);
            }
        }
        private void btn_upload_Click_3(object sender, RoutedEventArgs e)
        {
            try
            {
                string filePath = string.Empty;
                OpenFileDialog openFileDialog = new OpenFileDialog();
                if (openFileDialog.ShowDialog() == true)
                    filePath = openFileDialog.FileName;

                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    poitems_3 = PullIndentData(stream);
                }
                grid_quote_3.ItemsSource = null;
                grid_quote_3.ItemsSource = poitems_3;
                File.SetAttributes(filePath, FileAttributes.Normal);
                string targetPath = Convert.ToString(ConfigurationManager.AppSettings["TargetReportPath"]);

                targetPath = targetPath + "\\" + openFileDialog.SafeFileName;
                File.Move(filePath, targetPath, true);
                MessageBox.Show("The file has been processed and data has been uploaded.",
                        "Order Management System", MessageBoxButton.OK, MessageBoxImage.Information);
                // log.Info("The file has been processed and data has been uploaded.");
                //}
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to upload the data from the file. An error occured : " + ex.Message,
                    "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                // log.Error("Unable to upload the data from the file. An error occured : " + ex.Message);
            }
        }
        private List<Poitem> PullIndentData(FileStream stream)
        {
            List<Poitem> poitems = new List<Poitem>();
            try
            {
                List<List<string>> _readData = new List<List<string>>();
                bool isHeader = false;
                {
                    //string raisedBy = string.Empty;
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        do
                        {
                            while (reader.Read()) //Each ROW
                            {
                                List<string> lst = new List<string>();
                                for (int column = 0; column < reader.FieldCount; column++)
                                {
                                    //Console.WriteLine(reader.GetString(column));//Will blow up if the value is decimal etc. 
                                    if (reader.GetValue(column) != null)
                                        if (!Convert.ToString(reader.GetValue(column)).ToLower().Contains("description"))
                                        {
                                            lst.Add(Convert.ToString(reader.GetValue(column)));
                                            isHeader = false;
                                        }
                                        else
                                        {
                                            isHeader = true;
                                            break;
                                        }
                                }
                                if (!isHeader)
                                    if (lst.Count > 0)
                                        _readData.Add(lst);
                            }
                            foreach (List<string> lst in _readData)
                            {
                                if (lst.Count > 0)
                                {
                                    Poitem poitem = new Poitem();
                                    poitem.Sl_NO = Convert.ToInt32(lst[0]);
                                    poitem.Description = lst[1];
                                    poitem.Quantity = Convert.ToInt32(lst[2]);
                                    poitem.Units = lst[3];
                                    poitem.Unit_Price = Convert.ToDouble(lst[4]);
                                    poitem.Total_Price = Convert.ToDouble(lst[5]);

                                    poitems.Add(poitem);
                                    // loadExcelIndent.ExcelIndents = lstGridIndent;
                                }
                            }
                            //}
                        } while (reader.NextResult()); //Move to NEXT SHEET
                    }
                }
                return poitems;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to read the data from the file. An error occured : " + ex.Message,
                    "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                return poitems;
            }
        }


        private void grid_quote_3_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void grid_quote_2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void checkbox_Approve1_Checked(object sender, RoutedEventArgs e)
        {
            checkbox_Approve2.IsChecked = false;
            checkbox_Approve3.IsChecked = false;
            grid_po_confirmation.ItemsSource = null;
            grid_po_confirmation.ItemsSource = poitems_1;
        }
        private void checkbox_Approve2_Checked(object sender, RoutedEventArgs e)
        {
            checkbox_Approve1.IsChecked = false;
            checkbox_Approve3.IsChecked = false;
            grid_po_confirmation.ItemsSource = null;
            grid_po_confirmation.ItemsSource = poitems_2;
        }

        private void checkbox_Approve3_Checked(object sender, RoutedEventArgs e)
        {
            checkbox_Approve1.IsChecked = false;
            checkbox_Approve2.IsChecked = false;
            grid_po_confirmation.ItemsSource = null;
            grid_po_confirmation.ItemsSource = poitems_3;
        }

        private void grid_indentdata_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        protected override void OnClosing(CancelEventArgs e)
        {
            Menu menu = new Menu(_login);
            menu.Show();
        }

    }
}
