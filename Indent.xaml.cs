using ClosedXML.Excel;
using ExcelDataReader;
using Microsoft.Data.SqlClient;
using Microsoft.Win32;
using OrderManagementTool.Models;
using OrderManagementTool.Models.Excel;
using OrderManagementTool.Models.Indent;
using OrderManagementTool.Models.LogIn;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
//using ILog = log4net.ILog;
//using LogManager = log4net.LogManager;

namespace OrderManagementTool
{
    /// <summary>
    /// Interaction logic for Indent.xaml
    /// </summary>
    public partial class Indent : Window
    {
        //ILog log = LogManager.GetLogger(typeof(MainWindow));
        private string path = Convert.ToString(ConfigurationManager.AppSettings["InputFilePath"]);
        //private string targetPath = Convert.ToString(ConfigurationManager.AppSettings["TargetReportPath"]);
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

        public Indent(Login login)
        {
            try
            {
                _login = login;
                // mailFrom = _login.UserEmail;
                ////log.Info("In Indent Screen...");
                InitializeComponent();
                LoadItemCategoryName();
                LoadApprovalStatus();
                txt_raised_by.Text = _login.UserEmail;
                this.datepicker_date1.SelectedDate = DateTime.Now.Date;
                lbl_approval_status.Content = "Awaiting Approval";
                // lbl_approval_status.BorderBrush = System.Windows.Media.Brushes.Blue;
                lbl_approval_status.Background = System.Windows.Media.Brushes.Aqua;
                //lbl_approval_status.Foreground ();
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured during Load. " + ex.Message, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                ////log.Error("Error while fetching Indent information: " + ex.StackTrace);
            }
        }

        public Indent(Login login, long indentNo)
        {
            try
            {
                _login = login;
                ////log.Info("In Indent Screen...");
                InitializeComponent();
                LoadItemCategoryName();
                LoadApprovalStatus();
                txt_raised_by.Text = _login.UserEmail;
                txt_indent_no.Text = indentNo.ToString();
                GetIndent(indentNo);

                this.datepicker_date1.SelectedDate = DateTime.Now.Date;
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured during Indent Retrival. " + ex.Message, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                this.Close();
                ////log.Error("Error while fetching Indent information: " + ex.StackTrace);
            }
            //Image img = new Image();
            //img.Source = new BitmapImage(new Uri(@"~/Images/create-icon.png"));

            //StackPanel stackPnl = new StackPanel();
            //stackPnl.Orientation = Orientation.Horizontal;
            //stackPnl.Margin = new Thickness(10);
            //stackPnl.Children.Add(img);

            //Button btn = new Button();
            //btn.Content = stackPnl;
            //datepicker_date.Children.Add(btn);
        }
        private void DisableAllFields()
        {
            datepicker_date1.Focusable = false;
            datepicker_date1.IsEnabled = false;
            txt_indent_no.IsEnabled = false;
            txt_raised_by.IsEnabled = false;
            cbx_location_id.IsEnabled = false;
            cbx_approval_id.IsEnabled = false;
            cbx_itemcategoryname.IsEnabled = false;
            cbx_itemcode.IsEnabled = false;
            txt_item_name.IsEnabled = false;
            txt_description.IsEnabled = false;
            txt_quantity.IsEnabled = false;
            txt_technical_description.IsEnabled = false;
            cbx_units.IsEnabled = false;
            txt_remarks.IsEnabled = false;
            btn_add_category.IsEnabled = false;
            btn_add_itemcode.IsEnabled = false;
            btn_upload.IsEnabled = false;
            btn_update_field.IsEnabled = false;
            btn_clear_fields.IsEnabled = false;
            btn_remove_indent.IsEnabled = false;
            btn_generate_report.IsEnabled = false;
            btn_create_indent.IsEnabled = false;
            btn_generate_report.IsEnabled = false;
            grid_indentdata.IsReadOnly = true;
            txt_Revision_Remarks.IsEnabled = false;
            //grid_indentdata.IsEnabled = false;
            //btn_update_field.Focusable = false;
            //btn_clear_fields.IsEnabled = false;
            //btn_clear_fields.Focusable = false;
            //btn_remove_indent.IsEnabled = false;
            //btn_remove_indent.Focusable = false;
            //btn_add_indent.IsEnabled = false;
            //btn_add_indent.Focusable = false;
            //btn_create_indent.IsEnabled = false;
            //btn_create_indent.Focusable = false;
            //btn_generate_report.IsEnabled = false;
            //btn_generate_report.Focusable = false;
            //grid_indentdata.IsReadOnly = true;
            //grid_indentdata.Focusable = false;
        }
        private void GetIndent(long indentNo)
        {
            try
            {
                ////log.Info("Getting Indent infomration for Indent No: " + indentNo);
                using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlConnection"].ToString()))
                {
                    connection.Open();
                    //SqlCommand testCMD2 = new SqlCommand("GetEmployees", connection);
                    //testCMD2.CommandType = CommandType.StoredProcedure;
                    //testCMD2.Parameters.Add(new SqlParameter("@EmployeeId", System.Data.SqlDbType.BigInt, 50) { Value = _login.EmployeeID });
                    //List<string> userId = new List<string>();
                    //string userIdComma = "";


                    //using (SqlDataReader reader = testCMD2.ExecuteReader())
                    //{
                    //    while (reader.Read())
                    //    {
                    //        userId.Add(reader[0].ToString());
                    //    }
                    //    userIdComma = String.Join(",", userId);
                    //}
                    SqlCommand testCMD = new SqlCommand("GetIndent", connection);
                    testCMD.CommandType = CommandType.StoredProcedure;

                    testCMD.Parameters.Add(new SqlParameter("@IndentID", System.Data.SqlDbType.BigInt, 50) { Value = indentNo });
                    testCMD.Parameters.Add(new SqlParameter("@UserID", System.Data.SqlDbType.VarChar, 300) { Value = _login.EmployeeID });

                    // SqlDataReader dataReader = testCMD.ExecuteReader();
                    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(testCMD);

                    DataSet dataSet = new DataSet();
                    sqlDataAdapter.Fill(dataSet);

                    SaveIndent saveIndent = new SaveIndent();

                    int counter = 0;

                    while (counter < dataSet.Tables[0].Rows.Count)
                    {
                        saveIndent.Date = Convert.ToDateTime(dataSet.Tables[0].Rows[counter]["Date"]);
                        saveIndent.LocationCode = Convert.ToInt64(dataSet.Tables[0].Rows[counter]["LocationCode"]);
                        saveIndent.IndentRemarks = Convert.ToString(dataSet.Tables[0].Rows[counter]["Remarks"]);
                        saveIndent.ApproverName = Convert.ToString(dataSet.Tables[0].Rows[counter]["Approver"]);
                        saveIndent.ApprovalID = Convert.ToInt64(dataSet.Tables[0].Rows[counter]["Approver ID"]);
                        saveIndent.ApprovalStatus = Convert.ToString(dataSet.Tables[0].Rows[counter]["ApprovalStatus"]);

                        GridIndent gridIndent = new GridIndent();
                        gridIndent.SlNo = counter + 1;
                        gridIndent.ItemCategoryName = Convert.ToString(dataSet.Tables[0].Rows[counter]["ItemCategoryName"]);
                        gridIndent.ItemName = Convert.ToString(dataSet.Tables[0].Rows[counter]["ItemName"]);
                        gridIndent.ItemCode = Convert.ToString(dataSet.Tables[0].Rows[counter]["ItemCode"]);
                        gridIndent.Units = Convert.ToString(dataSet.Tables[0].Rows[counter]["Unit"]);
                        gridIndent.Description = Convert.ToString(dataSet.Tables[0].Rows[counter]["Description"]);
                        gridIndent.Technical_Specifications = Convert.ToString(dataSet.Tables[0].Rows[counter]["TechnicalSpecification"]);
                        gridIndent.Quantity = Convert.ToInt32(dataSet.Tables[0].Rows[counter]["Quantity"]);
                        gridIndent.Remarks = Convert.ToString(dataSet.Tables[0].Rows[counter]["Item Remarks"]);
                        gridIndents.Add(gridIndent);

                        saveIndent.Email = Convert.ToString(dataSet.Tables[0].Rows[counter]["Email"]);
                        counter++;
                    }

                    dataSet.Dispose();
                    txt_raised_by.Text = saveIndent.Email;
                    datepicker_date1.SelectedDate = saveIndent.Date;
                    LoadApprovalStatusRetrival();
                    LoadLocationId();
                    cbx_location_id.SelectedValue = saveIndent.LocationCode;
                    cbx_approval_id.SelectedValue = saveIndent.ApprovalID;

                    lbl_approval_status.Content = saveIndent.ApprovalStatus;
                    txt_Revision_Remarks.Text = saveIndent.IndentRemarks;
                    grid_indentdata.ItemsSource = null;
                    grid_indentdata.ItemsSource = gridIndents;
                    if (saveIndent.ApprovalStatus != "Enquiry Required")
                    {
                        DisableAllFields();
                        isGridReadOnly = true;
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured during Data Retrival. " + ex.Message, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                ////log.Error("Error while fetching Indent information: " + ex.StackTrace);
            }
        }

        private void DataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //DataRowView rowview = grid_indentdata.SelectedItem as DataRowView;
            try
            {
                if (!isGridReadOnly)
                {
                    var rowview = grid_indentdata.SelectedItem as GridIndent;
                    gridSelectedIndex = grid_indentdata.SelectedIndex;

                    if (rowview != null)
                    {
                        txt_quantity.Text = Convert.ToInt32(rowview.Quantity).ToString();
                        LoadItemCategoryName();
                        cbx_itemcategoryname.SelectedItem = rowview.ItemCategoryName;
                        LoadItemCode();
                        cbx_itemcode.SelectedItem = rowview.ItemCode;
                        LoadDescription();
                        LoadUnits();
                        cbx_units.SelectedItem = rowview.Units;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred :" + ex.Message,
                                   "Order Management System",
                                       MessageBoxButton.OK,
                                           MessageBoxImage.Error);
                ////log.Error("Error while selection of items in datagrid: " + ex.StackTrace);
            }
        }

        private void cbx_itemcategoryname_DropDownOpened(object sender, EventArgs e)
        {
            try
            {
                LoadItemCategoryName();
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred :" + ex.Message,
                                   "Order Management System",
                                       MessageBoxButton.OK,
                                           MessageBoxImage.Error);
                ////log.Error("Error while selecting item category: " + ex.StackTrace);
            }
        }

        private void cbx_itemcode_DropDownOpened(object sender, EventArgs e)
        {
            try
            {
                LoadItemCode();
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred :" + ex.Message,
                                   "Order Management System",
                                       MessageBoxButton.OK,
                                           MessageBoxImage.Error);
                ////log.Error("Error while selecting item code: " + ex.StackTrace);
            }
        }

        private void cbx_units_DropDownOpened(object sender, EventArgs e)
        {
            try
            {
                LoadUnits();
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred :" + ex.Message,
                                   "Order Management System",
                                       MessageBoxButton.OK,
                                           MessageBoxImage.Error);
                ////log.Error("Error while selecting units: " + ex.StackTrace);
            }
        }

        private void txt_quantity_lostfocus(object sender, RoutedEventArgs e)
        {
            if (Int32.TryParse(txt_quantity.Text, out int value))
            {
                quantityEntered = Convert.ToInt32(txt_quantity.Text);
            }
            else
            {
                MessageBox.Show("Please enter valid number",
                                   "Order Management System",
                                       MessageBoxButton.OK,
                                           MessageBoxImage.Error);
            }
        }

        private void cbx_itemcode_DropDownClosed(object sender, EventArgs e)
        {
            try
            {
                LoadDescription();
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred :" + ex.Message,
                                   "Order Management System",
                                       MessageBoxButton.OK,
                                           MessageBoxImage.Error);
                ////log.Error("Error while selecting Item code: " + ex.StackTrace);
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ////log.Info("Adding Indent information");
                if (cbx_itemcategoryname.SelectedItem != null && cbx_itemcode.SelectedItem != null && txt_quantity.Text != ""
                   && quantityEntered != 0 && cbx_units.SelectedItem != null)
                {
                    bool itemPresent = false;
                    GridIndent gridIndent = new GridIndent();
                    gridIndent.SlNo = gridIndents.Count + 1;
                    gridIndent.ItemName = txt_item_name.Text;
                    gridIndent.ItemCode = cbx_itemcode.SelectedItem.ToString();
                    gridIndent.ItemCategoryName = cbx_itemcategoryname.SelectedItem.ToString();
                    gridIndent.Quantity = quantityEntered;
                    gridIndent.Technical_Specifications = txt_technical_description.Text.Trim();
                    gridIndent.Units = cbx_units.SelectedItem.ToString();
                    gridIndent.Remarks = txt_remarks.Text.Trim();

                    foreach (var i in gridIndents)
                    {
                        if (i.ItemCategoryName == gridIndent.ItemCategoryName &&
                                i.ItemCode == gridIndent.ItemCode &&
                                    i.Description == gridIndent.Description &&
                                        i.Technical_Specifications == gridIndent.Technical_Specifications &&
                                            i.Units == gridIndent.Units && itemPresent == false)
                        {
                            itemPresent = true;

                        }
                    }
                    if (!itemPresent)
                    {
                        gridIndents.Add(gridIndent);
                        grid_indentdata.ItemsSource = null;
                        grid_indentdata.ItemsSource = gridIndents;
                        ClearFields();
                    }
                    else
                    {
                        MessageBox.Show("Item is already available!",
                            "Order Management System",
                                MessageBoxButton.OK,
                                    MessageBoxImage.Error);
                    }
                }
                else if (cbx_itemcategoryname.SelectedItem == null)
                {
                    MessageBox.Show("Please select Item Name",
                                        "Order Management System",
                                            MessageBoxButton.OK,
                                                MessageBoxImage.Error);
                }
                else if (cbx_itemcode.SelectedItem == null)
                {
                    MessageBox.Show("Please select Item Code",
                                        "Order Management System",
                                            MessageBoxButton.OK,
                                                MessageBoxImage.Error);
                }
                else if (txt_quantity.Text == "")
                {
                    MessageBox.Show("Please enter quantity",
                                        "Order Management System",
                                            MessageBoxButton.OK,
                                                MessageBoxImage.Error);
                }
                else if (quantityEntered == 0)
                {
                    MessageBox.Show("Please enter valid quantity",
                                        "Order Management System",
                                            MessageBoxButton.OK,
                                                MessageBoxImage.Error);
                }
                else if (cbx_units.SelectedItem == null)
                {
                    MessageBox.Show("Please select units",
                                        "Order Management System",
                                            MessageBoxButton.OK,
                                                MessageBoxImage.Error);
                }
                //////log.Error("Indent infomration added");
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred :" + ex.Message,
                                   "Order Management System",
                                       MessageBoxButton.OK,
                                           MessageBoxImage.Error);
                //////log.Error("Error while adding Indent information: " + ex.StackTrace);
            }
        }

        private void LoadApprovalStatus()
        {
            try
            {
                cbx_approval_id.SelectedValuePath = "Key";
                cbx_approval_id.DisplayMemberPath = "Value";
                ////log.Error("Loading Approval infomration");
                cbx_approval_id.Items.Clear();
                var exceptionList = new List<string> { "Clerk", "Supervisor" };
                //var data = (from emp in orderManagementContext.Employee
                //            join des in orderManagementContext.Designation
                //            on emp.DesignationId equals des.DesignationId
                //            where emp.EmployeeId != _login.EmployeeID
                //            where emp.DesignationId < (from emp1 in orderManagementContext.Employee where emp1.EmployeeId == _login.EmployeeID select emp1.DesignationId).FirstOrDefault()
                //            select
                //            new
                //            {
                //                emp.FirstName,
                //                emp.LastName,
                //                des.Designation1,
                //                emp.EmployeeId
                //            })
                //                 .Distinct().ToList();
                //foreach (string s in exceptionList)
                //{
                //    var remove = data.Select(e => e.Designation1 == s).ToList();
                //    int index = remove.IndexOf(true);
                //    if (index >= 0)
                //        data.RemoveAt(index);
                //}

                var data = (from emp in orderManagementContext.Employee
                            where emp.EmployeeId == (from emp1 in orderManagementContext.Employee
                                                     where emp1.EmployeeId == _login.EmployeeID
                                                     select emp1.ReportsTo).FirstOrDefault()
                            select new
                            {
                                emp.FirstName,
                                emp.LastName,
                                emp.EmployeeId
                            });


                foreach (var i in data)
                {
                    cbx_approval_id.Items.Add(new KeyValuePair<long, string>(i.EmployeeId, i.FirstName.Trim() + " " + i.LastName.Trim()));
                }
                ////log.Error("Approval Information loaded.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured during Approval ID fetch " + ex.Message, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                ////log.Error("Error while loading approval : " + ex.StackTrace);
            }
        }

        private void LoadApprovalStatusRetrival()
        {
            try
            {
                cbx_approval_id.SelectedValuePath = "Key";
                cbx_approval_id.DisplayMemberPath = "Value";
                ////log.Error("Loading Approval infomration");
                cbx_approval_id.Items.Clear();

                var data = (from emp in orderManagementContext.Employee
                            select new
                            {
                                emp.FirstName,
                                emp.LastName,
                                emp.EmployeeId
                            });


                foreach (var i in data)
                {
                    cbx_approval_id.Items.Add(new KeyValuePair<long, string>(i.EmployeeId, i.FirstName.Trim() + " " + i.LastName.Trim()));
                }
                ////log.Error("Approval Information loaded.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured during Retrival Approval ID fetch " + ex.Message, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                ////log.Error("Error while loading approval : " + ex.StackTrace);
            }
        }
        private void LoadLocationId()
        {
            try
            {
                cbx_location_id.SelectedValuePath = "Key";
                cbx_location_id.DisplayMemberPath = "Value";
                ////log.Error("Loading Approval infomration");
                cbx_location_id.Items.Clear();

                var data = (from loc in orderManagementContext.LocationCode
                            select loc).ToList();


                foreach (var i in data)
                {
                    cbx_location_id.Items.Add(new KeyValuePair<long, string>(i.LocationCodeId, i.LocationId + " - " + i.LocationName.Trim()));
                }
                ////log.Error("Approval Information loaded.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured during Location Code fetch " + ex.Message, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                ////log.Error("Error while loading location code : " + ex.StackTrace);
            }
        }

        private void LoadItemCategoryName()
        {
            cbx_itemcategoryname.Items.Clear();
            var itemCategoryName = (from i in orderManagementContext.ItemCategory
                                    select i.ItemCategoryName).Distinct().ToList();
            foreach (var i in itemCategoryName)
            {
                cbx_itemcategoryname.Items.Add(i.Trim());
            }
        }
        private void LoadDescription()
        {
            selectedItemCode = cbx_itemcode.SelectionBoxItem.ToString();
            var itemDetails = (from i in orderManagementContext.ItemMaster
                               from ic in orderManagementContext.ItemCategory
                               where i.ItemCategoryId == ic.ItemCategoryId &&
                               ic.ItemCategoryName == selectedItemCategoryName &&
                               i.ItemCode == selectedItemCode
                               select i).FirstOrDefault();
            if (itemDetails != null)
            {
                txt_description.Text = itemDetails.Description;
                txt_technical_description.Text = itemDetails.TechnicalSpecification;
                txt_item_name.Text = itemDetails.ItemName;
            }
        }
        private void LoadUnits()
        {
            cbx_units.Items.Clear();
            units = (from u in orderManagementContext.UnitMaster
                     select u.Unit).Distinct().ToList();
            foreach (var u in units)
            {
                cbx_units.Items.Add(u.Trim());
            }
        }
        private void LoadItemCode()
        {
            if (cbx_itemcategoryname.SelectedItem != null)
            {
                selectedItemCategoryName = cbx_itemcategoryname.SelectedItem.ToString();
                cbx_itemcode.Items.Clear();
                itemCode = (from i in orderManagementContext.ItemMaster
                            from ic in orderManagementContext.ItemCategory
                            where i.ItemCategoryId == ic.ItemCategoryId && ic.ItemCategoryName == selectedItemCategoryName
                            select i.ItemCode).Distinct().ToList();
                foreach (var i in itemCode)
                {
                    cbx_itemcode.Items.Add(i);
                }
            }
            else
            {
                MessageBox.Show("Please select item name", "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void ClearFields()
        {
            cbx_itemcategoryname.SelectedValue = "";
            cbx_itemcode.SelectedValue = "";
            txt_description.Text = "";
            txt_quantity.Text = "";
            txt_technical_description.Text = "";
            cbx_units.SelectedValue = "";
            txt_remarks.Text = "";
            txt_technical_description.Text = "";
            txt_item_name.Text = "";
        }

        //private void LoadApprovalId()
        //{
        //    LoadApprovalStatus();
        //}

        private void btn_clear_fields_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                ////log.Info("Clearing Fields");
                ClearFields();
                ////log.Info("Cleared fields");
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred :" + ex.Message,
                                   "Order Management System",
                                       MessageBoxButton.OK,
                                           MessageBoxImage.Error);
                ////log.Error("Error while clearing Indent information: " + ex.StackTrace);
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            ////log.Info("Adding data to grid");
            if (gridSelectedIndex >= 0)
            {
                // bool itemPresent = false;
                GridIndent value = new GridIndent();
                value.ItemCode = cbx_itemcode.SelectedItem.ToString();
                value.ItemCategoryName = cbx_itemcategoryname.SelectedItem.ToString();
                value.ItemName = txt_item_name.Text;
                value.Quantity = quantityEntered;
                value.Technical_Specifications = txt_technical_description.Text.Trim();
                value.Units = cbx_units.SelectedItem.ToString();
                value.Remarks = txt_remarks.Text.Trim();

                //foreach (var i in gridIndents)
                //{
                //    if (i.ItemName == value.ItemName && i.ItemCode == value.ItemCode && i.Description == value.Description && i.Technical_Specifications == value.Technical_Specifications && itemPresent == false)
                //    {
                //        itemPresent = true;
                //    }
                //}
                //if (!itemPresent)
                //{
                var gridIndent = gridIndents[gridSelectedIndex];
                gridIndent.ItemCode = cbx_itemcode.SelectedItem.ToString();
                gridIndent.ItemCategoryName = cbx_itemcategoryname.SelectedItem.ToString();
                gridIndent.Quantity = quantityEntered;
                gridIndent.ItemName = txt_item_name.Text;
                gridIndent.Technical_Specifications = txt_technical_description.Text.Trim();
                gridIndent.Units = cbx_units.SelectedItem.ToString();
                gridIndent.Remarks = txt_remarks.Text.Trim();
                grid_indentdata.ItemsSource = null;
                grid_indentdata.ItemsSource = gridIndents;
                gridSelectedIndex = -1;
                ClearFields();
                // }
                ////log.Info("Data added to grid");
            }
            else
            {
                MessageBox.Show("Item Already Present");
            }
            // }
        }

        private void txt_quantity_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void cbx_approval_id_DropDownOpened(object sender, EventArgs e)
        {
            try
            {
                LoadApprovalStatus();
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred :" + ex.Message,
                                   "Order Management System",
                                       MessageBoxButton.OK,
                                           MessageBoxImage.Error);
                ////log.Error("Error while loading approval : " + ex.StackTrace);
            }
        }

        private void btn_create_indent_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                this.Cursor = Cursors.Wait;
                ////log.Info("Creating Indent...");
                if (datepicker_date1.SelectedDate.ToString() == "")
                {
                    MessageBox.Show("Please enter Valid Date");
                    this.Cursor = null;
                    return;
                }
                if (cbx_approval_id.SelectedValue == null)
                {
                    MessageBox.Show("Please select Approver");
                    this.Cursor = null;
                    return;
                }
                if (cbx_approval_id.SelectedValue == null)
                {
                    MessageBox.Show("Please Enter Location");
                    this.Cursor = null;
                    return;
                }

                if (gridIndents.Count == 0)
                {
                    MessageBox.Show("Please Enter Indent Details");
                    this.Cursor = null;
                    return;
                }

                else
                {
                    SaveIndent saveIndent = new SaveIndent();
                    string selected_date;


                    if (datepicker_date1.SelectedDate.ToString().Contains("-"))
                    {
                        var fulldate = datepicker_date1.SelectedDate.ToString().Split(" ");
                        var date = fulldate[0].Split("-");
                        //var date = fulldate[0].Replace("-", "/");

                        var modifiedDate = date[2] + "/" + date[1] + "/" + date[0];
                        // DateTime enteredDate = Convert.ChangeType(modifiedDate, typeof(DateTime));
                        DateTime enteredDate = new DateTime(Convert.ToInt32(date[2]), Convert.ToInt32(date[1]), Convert.ToInt32(date[0]));
                        // DateTime enteredDate = modifiedDate.ToDateTime("yyyy/d/M HH:mm");

                        //CultureInfo culture = new CultureInfo("en-US");
                        //DateTime tempDate = Convert.ToDateTime(date, culture);

                        CultureInfo provider = CultureInfo.InvariantCulture;
                        //DateTime dateTime16 = DateTime.ParseExact(date, new string[] { "MM.dd.yyyy", "MM-dd-yyyy", "MM/dd/yyyy" }, provider, DateTimeStyles.None);

                        //selected_date = DateTime.ParseExact(date, "dd/MM/yyyy", CultureInfo.InvariantCulture)
                        //    .ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                    }
                    saveIndent.Date = datepicker_date1.SelectedDate;
                    saveIndent.LocationCode = Convert.ToInt64(cbx_location_id.SelectedValue);
                    saveIndent.RaisedBy = _login.EmployeeID;
                    saveIndent.CreateDate = DateTime.Now;
                    // var approvalID = (from a in orderManagementContext.UserMaster
                    //                  where a.Email == cbx_approval_id.SelectedValue.ToString()
                    //                  select a.UserId).FirstOrDefault();
                    saveIndent.ApprovalID = Convert.ToInt64(cbx_approval_id.SelectedValue);
                    // saveIndent.ApprovalID = 1;
                    saveIndent.GridIndents = gridIndents;
                    bool isMailSent = false;

                    using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlConnection"].ToString()))
                    {
                        if (txt_indent_no.Text == null || txt_indent_no.Text == "")
                        {
                            connection.Open();
                            SqlCommand testCMD = new SqlCommand("create_indent", connection);
                            testCMD.CommandType = CommandType.StoredProcedure;

                            //testCMD.Parameters.Add(new SqlParameter("@Date", System.Data.SqlDbType.VarChar, 50) { Value = saveIndent.Date });
                            testCMD.Parameters.Add(new SqlParameter("@LocationCode", System.Data.SqlDbType.BigInt, 50) { Value = saveIndent.LocationCode });
                            testCMD.Parameters.Add(new SqlParameter("@RaisedBy", System.Data.SqlDbType.BigInt, 50) { Value = saveIndent.RaisedBy });
                            //testCMD.Parameters.Add(new SqlParameter("@CreatedDate", System.Data.SqlDbType.DateTime, 50) { Value = saveIndent.CreateDate });
                            testCMD.Parameters.Add(new SqlParameter("@IndentId", System.Data.SqlDbType.BigInt, 50) { Value = saveIndent.IndentId });
                            testCMD.Parameters["@IndentId"].Direction = ParameterDirection.Output;

                            testCMD.ExecuteNonQuery(); // read output value from @NewId 
                            saveIndent.IndentId = Convert.ToInt32(testCMD.Parameters["@IndentId"].Value);

                            indentNo = Convert.ToInt32(testCMD.Parameters["@IndentId"].Value);

                            foreach (var i in saveIndent.GridIndents)
                            {
                                SqlCommand testCMD1 = new SqlCommand("create_indentDetails", connection);
                                testCMD1.CommandType = CommandType.StoredProcedure;
                                testCMD1.Parameters.Add(new SqlParameter("@IndentID", System.Data.SqlDbType.BigInt, 50) { Value = saveIndent.IndentId });
                                testCMD1.Parameters.Add(new SqlParameter("@ItemName", System.Data.SqlDbType.VarChar, 50) { Value = i.ItemCategoryName });
                                testCMD1.Parameters.Add(new SqlParameter("@ItemCode", System.Data.SqlDbType.VarChar, 50) { Value = i.ItemCode });
                                testCMD1.Parameters.Add(new SqlParameter("@Unit", System.Data.SqlDbType.VarChar, 50) { Value = i.Units });
                                testCMD1.Parameters.Add(new SqlParameter("@Quantity", System.Data.SqlDbType.VarChar, 50) { Value = i.Quantity });
                                //testCMD1.Parameters.Add(new SqlParameter("@CreatedDate", System.Data.SqlDbType.VarChar, 50) { Value = saveIndent.CreateDate });
                                testCMD1.Parameters.Add(new SqlParameter("@CreatedBy", System.Data.SqlDbType.BigInt, 50) { Value = saveIndent.RaisedBy });
                                testCMD1.Parameters.Add(new SqlParameter("@Remarks", System.Data.SqlDbType.VarChar, 50) { Value = i.Remarks });
                                testCMD1.ExecuteNonQuery(); // read output value from @NewId 
                            }

                            SqlCommand testCMD2 = new SqlCommand("create_indent_Approval", connection);
                            testCMD2.CommandType = CommandType.StoredProcedure;
                            testCMD2.Parameters.Add(new SqlParameter("@IndentID", System.Data.SqlDbType.BigInt, 50) { Value = saveIndent.IndentId });
                            testCMD2.Parameters.Add(new SqlParameter("@ApprovalID", System.Data.SqlDbType.BigInt, 50) { Value = saveIndent.ApprovalID });
                            testCMD2.Parameters.Add(new SqlParameter("@ApprovalStatusID", System.Data.SqlDbType.BigInt, 50) { Value = 1 });
                            // testCMD2.Parameters.Add(new SqlParameter("@CreatedDate", System.Data.SqlDbType.DateTime, 50) { Value = saveIndent.CreateDate });
                            testCMD2.Parameters.Add(new SqlParameter("@CreatedBy", System.Data.SqlDbType.BigInt, 50) { Value = saveIndent.RaisedBy });
                            testCMD2.ExecuteNonQuery();
                            mailTo = (from a in orderManagementContext.UserMaster
                                      where a.EmployeeId == saveIndent.ApprovalID
                                      select a.Email).FirstOrDefault();
                            txt_indent_no.Text = indentNo.ToString();
                            string body = GenerateIndent(indentNo.ToString());
                            isMailSent = SendMail("Indent Number " + indentNo + " is updated by " + _login.UserName, body);
                            if (isMailSent)
                            {
                                MessageBox.Show("The Indent  " + indentNo + " is created.", "Order Management System",
                      MessageBoxButton.OK, MessageBoxImage.Information);
                            }
                            else
                            {
                                MessageBox.Show("The Indent  " + indentNo + " is created. But mail is not sent", "Order Management System",
                     MessageBoxButton.OK, MessageBoxImage.Information);
                            }
                        }

                        else
                        {
                            connection.Open();
                            indentNo = Convert.ToInt64(txt_indent_no.Text);
                            SqlCommand testCMD = new SqlCommand("delete_indent_details", connection);
                            testCMD.Parameters.Add(new SqlParameter("@IndentID", System.Data.SqlDbType.BigInt, 50) { Value = indentNo });
                            testCMD.CommandType = CommandType.StoredProcedure;

                            testCMD.ExecuteNonQuery(); // read output value from @NewId 
                            saveIndent.IndentId = Convert.ToInt32(testCMD.Parameters["@IndentId"].Value);

                            foreach (var i in saveIndent.GridIndents)
                            {
                                string createDate = saveIndent.CreateDate.Month + "/" + saveIndent.CreateDate.Day + "/" + saveIndent.CreateDate.Year;
                                SqlCommand testCMD1 = new SqlCommand("create_indentDetails", connection);
                                testCMD1.CommandType = CommandType.StoredProcedure;
                                testCMD1.Parameters.Add(new SqlParameter("@IndentID", System.Data.SqlDbType.BigInt, 50) { Value = saveIndent.IndentId });
                                testCMD1.Parameters.Add(new SqlParameter("@ItemName", System.Data.SqlDbType.VarChar, 50) { Value = i.ItemCategoryName });
                                testCMD1.Parameters.Add(new SqlParameter("@ItemCode", System.Data.SqlDbType.VarChar, 50) { Value = i.ItemCode });
                                testCMD1.Parameters.Add(new SqlParameter("@Unit", System.Data.SqlDbType.VarChar, 50) { Value = i.Units });
                                testCMD1.Parameters.Add(new SqlParameter("@Quantity", System.Data.SqlDbType.VarChar, 50) { Value = i.Quantity });
                                //testCMD1.Parameters.Add(new SqlParameter("@CreatedDate", System.Data.SqlDbType.DateTime, 50) { Value = createDate });
                                testCMD1.Parameters.Add(new SqlParameter("@CreatedBy", System.Data.SqlDbType.BigInt, 50) { Value = saveIndent.RaisedBy });
                                testCMD1.Parameters.Add(new SqlParameter("@Remarks", System.Data.SqlDbType.VarChar, 50) { Value = i.Remarks });
                                testCMD1.ExecuteNonQuery(); // read output value from @NewId 
                            }

                            SqlCommand testCMD2 = new SqlCommand("update_indent_Approval", connection);
                            testCMD2.CommandType = CommandType.StoredProcedure;
                            testCMD2.Parameters.Add(new SqlParameter("@IndentID", System.Data.SqlDbType.BigInt, 50) { Value = saveIndent.IndentId });
                            testCMD2.Parameters.Add(new SqlParameter("@ApprovalID", System.Data.SqlDbType.BigInt, 50) { Value = saveIndent.ApprovalID });
                            testCMD2.Parameters.Add(new SqlParameter("@ApprovalStatusID", System.Data.SqlDbType.BigInt, 50) { Value = 1 });
                            //testCMD2.Parameters.Add(new SqlParameter("@CreatedDate", System.Data.SqlDbType.DateTime, 50) { Value = saveIndent.CreateDate });
                            testCMD2.Parameters.Add(new SqlParameter("@CreatedBy", System.Data.SqlDbType.BigInt, 50) { Value = saveIndent.RaisedBy });
                            testCMD2.ExecuteNonQuery();
                            mailTo = (from a in orderManagementContext.UserMaster
                                      where a.EmployeeId == saveIndent.ApprovalID
                                      select a.Email).FirstOrDefault();
                            string body = GenerateIndent(indentNo.ToString());
                            isMailSent = SendMail("Indent Number " + indentNo + " is updated by " + _login.UserName, body);

                            if (isMailSent)
                            {

                                MessageBox.Show("The Indent number " + indentNo + " has been updated.", "Order Management System",
                                   MessageBoxButton.OK, MessageBoxImage.Information);
                            }
                            else
                            {
                                MessageBox.Show("The Indent  " + indentNo + " is updated. But mail is not sent", "Order Management System",
                     MessageBoxButton.OK, MessageBoxImage.Information);
                            }
                        }
                        //log.Info("Indent created and generated.");

                        ////log.Info("Indent created and generated.");
                        DisableAllFields();
                        grid_indentdata.IsEnabled = false;
                        this.Cursor = null;
                    }
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show("An error occured during save. " + ex.Message, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                ////log.Error("Error while creating approval : " + ex.StackTrace);
            }
            finally
            {
                this.Cursor = null;
            }
        }

        private bool SendMail(string subject, string body)
        {
            bool mailSent = false;
            try
            {
                ////log.Info("Sending Mail..");

                string message = DateTime.Now + " In SendMail\n";

                using (MailMessage mm = new MailMessage())
                {
                    mm.From = new MailAddress(Convert.ToString(ConfigurationManager.AppSettings["MailFrom"]));
                    //string[] _toAddress = Convert.ToString(ConfigurationManager.AppSettings["MailTo"]).Split(';');
                    //foreach (string address in _toAddress)
                    //{
                    //    mm.To.Add(address);
                    //}

                    // mm.To.Add(mailFrom);
                    mm.To.Add(mailTo);

                    mm.Subject = "Indent Number - " + indentNo.ToString();
                    body = body + ". Please click on the link to approve or deny the indent. " + ConfigurationManager.AppSettings["url"] + "/" + indentNo;
                    mm.Body = body;
                    mm.IsBodyHtml = true;

                    // mm.Attachments.Add(new System.Net.Mail.Attachment(filePathLocation));

                    SmtpClient smtp = new SmtpClient();
                    smtp.Host = ConfigurationManager.AppSettings["Host"];
                    smtp.EnableSsl = false;
                    //smtp.EnableSsl = false;
                    NetworkCredential NetworkCred = new NetworkCredential(ConfigurationManager.AppSettings["Username"],
                        ConfigurationManager.AppSettings["Password"]);
                    smtp.UseDefaultCredentials = false;
                    smtp.Credentials = NetworkCred;
                    smtp.Port = int.Parse(ConfigurationManager.AppSettings["Port"]);

                    message = DateTime.Now + " Sending Mail\n";
                    smtp.Send(mm);
                    message = DateTime.Now + " Mail Sent\n";

                    System.Threading.Thread.Sleep(3000);
                    mailSent = true;
                }
                return mailSent;
                ////log.Info("Mail sent..");
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error while sending mail : " + ex.Message, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                ////log.Error("Error while Sending Mail : " + ex.StackTrace);
                return mailSent;
            }
        }

        private void btn_generate_report_Click(object sender, RoutedEventArgs e)
        {
            GenerateIndent(txt_indent_no.Text);
        }
        private string GenerateIndent(string indentNo)
        {
            Dictionary<string, string> _headers = GetHeaders();

            IndentHeader headerData = new IndentHeader();
            headerData.To = "M/S. GILIYAL INDUSTRIES";
            headerData.AddressLine1 = "AddressLine1";
            headerData.AddressLine2 = "AddressLine2";
            headerData.AddressLine3 = "AddressLine3";
            headerData.Contact = "Contact";
            headerData.GSTNo = "GST No";
            headerData.IECNo = "IEC No";
            headerData.IndentNo = "Indent No";
            headerData.IndentDate = DateTime.Now;
            headerData.RefNo = "Ref No";
            headerData.RefDate = DateTime.Now;
            headerData.Remarks = "Remarks";
            headerData.Attention = "Attention";
            headerData.IndentDate = datepicker_date1.SelectedDate.Value;
            headerData.Project = "Project";
            headerData.WBS = "WBS 1939";

            //GenerateIndent(_headers, headerData, gridIndents);
            string htmlText = GenerateIndentMail(_headers, headerData, gridIndents, indentNo);
            return htmlText;
        }
        private Dictionary<string, string> GetHeaders()
        {
            ////log.Info("Reading Header information from AppConfig");
            Dictionary<string, string> headers = new Dictionary<string, string>();
            var configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            foreach (var sectionKey in configuration.Sections.Keys)
            {
                var section = configuration.GetSection(sectionKey.ToString());
                var appSettings = section as AppSettingsSection;
                if (appSettings == null) continue;
                foreach (var key in appSettings.Settings.AllKeys)
                {
                    headers.Add(key, appSettings.Settings[key].Value);
                }
            }
            return headers;
        }

        private string GenerateIndentMail(Dictionary<string, string> headersAndFooters,
            IndentHeader headerData, List<GridIndent> gridIndents, string indentNo)
        {
            try
            {
                StringBuilder htmlString = new StringBuilder();
                htmlString.Append("<html>");
                htmlString.Append("<head>");
                htmlString.Append("<style>");
                htmlString.Append(".tableborder");
                htmlString.Append("{");
                htmlString.Append("border:1px;");
                htmlString.Append("border-style:solid;");
                htmlString.Append("border.color:red;");
                htmlString.Append("}");
                htmlString.Append(".tdDist {");
                htmlString.Append("width: 100%;");
                htmlString.Append("}");
                htmlString.Append("</style>");
                htmlString.Append("</head>");
                htmlString.Append("<body>");
                htmlString.Append("<table width='60%' border=1 border-style='solid' border-color='#8CBD48'>");
                htmlString.Append("<tr>");
                htmlString.Append("<td colspan=9 style='background-color:#EBF8F0;color:#6C6C6C;font-family:Verdana; vertical-align:middle'><h2><i><center>SLPP RENEW LLP<center></i></h2></td>");
                htmlString.Append("</tr>");
                htmlString.Append("<tr>");
                htmlString.Append("<td colspan=9 style='background-color:#EBF8F0;color:#6C6C6C;font-family:Verdana;vertical-align:middle'><h2><center><i><u>MATERIAL INDENT FORM</u></i></center></h2></td>");
                htmlString.Append("</tr>");
                htmlString.Append("<tr>");
                htmlString.Append("<td colspan=7 style='background-color:#EBF8F0;color:#6C6C6C;font-family:Verdana;'><i>Related WBS No: </i></td>");
                htmlString.Append("<td  style='background-color:#EBF8F0;color:#6C6C6C;font-family:Verdana;'><i>Indent No :</td>");
                htmlString.Append("<td  style='background-color:#EBF8F0;color:#6C6C6C;font-family:Verdana;'><i>" + indentNo + "</i></td></tr>");
                htmlString.Append("<tr>");
                htmlString.Append("<td colspan=7  style='background-color:#EBF8F0;color:#6C6C6C;font-family:Verdana;'><i>Approved By: </i></td>");
                htmlString.Append("<td style='background-color:#EBF8F0;color:#6C6C6C;font-family:Verdana;'><i>Approved On :</td>");
                htmlString.Append("<td style='background-color:#EBF8F0;color:#6C6C6C;font-family:Verdana;'><i></i></td></tr>");
                htmlString.Append("</table>");
                htmlString.Append("<table width='60%'  border=1 border-style='solid' border-color='#8CBD48'>");
                htmlString.Append("<tr style='background-color:#609F19;color:#FFFFFF;font-family:Verdana;font-size:11;vertical-align:middle'><td>Sl.No</td>");
                htmlString.Append("<td>Item Code</td>");
                htmlString.Append("<td>Item Category Name</td>");
                htmlString.Append("<td>Item Name</td>");
                htmlString.Append("<td>Technical Specifications</td>");
                htmlString.Append("<td>Units</td>");
                htmlString.Append("<td>Quantity</td>");
                htmlString.Append("<td>Remarks</td>");
                htmlString.Append("</tr>");
                int counter = 0;
                for (int k = 0; k < gridIndents.Count + 10; k++)
                {
                    if (k <= gridIndents.Count - 1)
                    {
                        htmlString.Append("<tr style='background-color:#EBF8F0;color:#6C6C6C;font-family:Verdana;font-size:11;'>");
                        htmlString.Append("<td>" + counter + 1 + "</td>");
                        htmlString.Append("<td>" + gridIndents[k].ItemCode + "</td>");
                        htmlString.Append("<td>" + gridIndents[k].ItemCategoryName + "</td>");
                        htmlString.Append("<td>" + gridIndents[k].ItemName + "</td>");
                        htmlString.Append("<td>" + gridIndents[k].Technical_Specifications + "</td>");
                        htmlString.Append("<td>" + gridIndents[k].Units + "</td>");
                        htmlString.Append("<td>" + gridIndents[k].Quantity + "</td>");
                        htmlString.Append("<td>" + gridIndents[k].Remarks + "</td>");
                        htmlString.Append("</tr>");
                    }
                }
                htmlString.Append("<tr>");
                htmlString.Append("<td colspan=9 style='background-color:#EBF8F0;color:#6C6C6C;font-family:Verdana;font-size:11;vertical-align:middle'><i>Revision Remarks: </i></td></tr>");
                htmlString.Append("<tr><td colspan=9 style='background-color:#EBF8F0;color:#6C6C6C;font-family:Verdana;font-size:11;vertical-align:middle'><i>1.          </i></td></tr>");
                htmlString.Append("<tr><td colspan=9 style='background-color:#EBF8F0;color:#6C6C6C;font-family:Verdana;font-size:11;vertical-align:middle'><i>2.          </i></td></tr>");
                htmlString.Append("</table>");
                return htmlString.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred. Please contact the administrator", "Order Management System",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }

        }

        private void GenerateIndent(Dictionary<string, string> headersAndFooters,
            IndentHeader headerData, List<GridIndent> gridIndents)
        {
            ////log.Info("Generating Indent Report..");
            try
            {
                var workBook = new XLWorkbook();
                workBook.AddWorksheet("Report");
                var worksheet = workBook.Worksheet("Report");

                //var imagePath = @"D:\Dinesh\Projects\GitHub\OrderManagementTool\Images\logo.png";
                //  var imagePath = @".\Images\logo.png";
                //var image = worksheet.AddPicture(imagePath)
                //    .MoveTo(worksheet.Cell("A1"))
                //    .Scale(.1).Placement = ClosedXML.Excel.Drawings.XLPicturePlacement.Move;
                var rangeMerged = worksheet.Range("A1:I1").Merge();
                rangeMerged.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                rangeMerged.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                rangeMerged.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                rangeMerged.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                rangeMerged.Style.Font.Bold = true;
                rangeMerged.Style.Font.FontSize = 16;
                rangeMerged.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                rangeMerged.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                var row = worksheet.Row(1);
                rangeMerged.Style.Fill.BackgroundColor = XLColor.LightYellow;
                rangeMerged.Style.Font.Italic = true;
                row.Height = 40;
                var rangeMerged1 = worksheet.Range("A2:I3").Merge();
                rangeMerged1.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                rangeMerged1.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                rangeMerged1.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                rangeMerged1.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                rangeMerged1.Style.Font.Bold = true;
                rangeMerged1.Style.Font.Italic = true;
                rangeMerged1.Style.Font.Underline = XLFontUnderlineValues.Single;
                rangeMerged1.Style.Font.FontSize = 16;
                rangeMerged1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                rangeMerged1.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                rangeMerged1.Style.Fill.BackgroundColor = XLColor.LightYellow1;
                var rangeMerged2 = worksheet.Range("A4:G4").Merge();
                rangeMerged2.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                rangeMerged2.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                rangeMerged2.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                rangeMerged2.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                rangeMerged2.Style.Font.Bold = true;
                rangeMerged2.Style.Font.Italic = true;
                rangeMerged2.Style.Font.FontSize = 12;

                var rangeMerged2a = worksheet.Range("A5:G5").Merge();
                rangeMerged2a.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                rangeMerged2a.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                rangeMerged2a.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                rangeMerged2a.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                rangeMerged2a.Style.Font.Bold = true;
                rangeMerged2a.Style.Font.Italic = true;
                rangeMerged2a.Style.Font.FontSize = 12;

                worksheet.Cell("A1").Value = worksheet.Cell("A1").Value + headersAndFooters["CompanyHeader"];
                worksheet.Cell("A2").Value = headersAndFooters["IndentHeader"];
                var rHeader = worksheet.Range("A6:I6");
                rHeader.Style.Fill.BackgroundColor = XLColor.Aqua;
                var rHeader1 = worksheet.Range("B6:I6");
                worksheet.Cell("A4").Value = headersAndFooters["Project"] + " " + headerData.Project;
                worksheet.Cell("A4").Style.Font.Bold = true;
                worksheet.Cell("A4").Style.Font.Italic = true;
                rHeader1.Style.Font.Bold = true;
                rHeader1.Style.Font.Italic = true;
                worksheet.Cell("H4").Value = headersAndFooters["IndentDate"];
                worksheet.Cell("H4").Style.Font.Bold = true;
                worksheet.Cell("H4").Style.Font.Italic = true;
                worksheet.Cell("I4").Value = headerData.IndentDate;
                worksheet.Cell("I4").Style.Font.Bold = true;
                worksheet.Cell("I4").Style.Font.Italic = true;

                worksheet.Cell("A5").Value = headersAndFooters["WBS"] + " " + headerData.WBS;
                worksheet.Cell("A5").Style.Font.Bold = true;
                worksheet.Cell("A5").Style.Font.Italic = true;
                worksheet.Cell("H5").Value = headersAndFooters["IndentNo"];
                worksheet.Cell("H5").Style.Font.Bold = true;
                worksheet.Cell("H5").Style.Font.Italic = true;
                worksheet.Cell("I5").Value = headerData.IndentNo;
                worksheet.Cell("I5").Style.Font.Bold = true;
                worksheet.Cell("I5").Style.Font.Italic = true;

                worksheet.Cell("B6").Value = "Related Item No.";
                worksheet.Cell("C6").Value = "Material Code";
                worksheet.Cell("D6").Value = "Unit";
                worksheet.Cell("E6").Value = "Total Planned";
                worksheet.Cell("F6").Value = "Total Supplied";
                worksheet.Cell("G6").Value = "Stock As On";
                worksheet.Cell("H6").Value = "Quantity Intended";
                worksheet.Cell("I6").Value = "Description / Specification";

                int j = 7;
                int i = 1;
                for (int k = 0; k < gridIndents.Count + 10; k++)
                {
                    if (k <= gridIndents.Count - 1)
                    {
                        worksheet.Cell("B" + j).Value = i;
                        worksheet.Cell("C" + j).Value = gridIndents[k].ItemCode;
                        worksheet.Cell("D" + j).Value = gridIndents[k].Units;
                        //worksheet.Cell("E" + j).Value = gridIndents[k].TotalPlanned;
                        //worksheet.Cell("F" + j).Value = gridIndents[k].TotalSupplied;
                        //worksheet.Cell("G" + j).Value = gridIndents[k].StockAsOn;
                        worksheet.Cell("E" + j).Value = 0;
                        worksheet.Cell("F" + j).Value = 0;
                        worksheet.Cell("G" + j).Value = 0;
                        worksheet.Cell("H" + j).Value = gridIndents[k].Quantity;
                        worksheet.Cell("I" + j).Value = gridIndents[k].Description;
                        j++;
                        i++;
                    }
                    else
                    {
                        worksheet.Cell("B" + j).Value = "";
                        worksheet.Cell("C" + j).Value = "";
                        worksheet.Cell("D" + j).Value = "";
                        worksheet.Cell("E" + j).Value = "";
                        worksheet.Cell("F" + j).Value = "";
                        worksheet.Cell("G" + j).Value = "";
                        worksheet.Cell("H" + j).Value = "";
                        worksheet.Cell("I" + j).Value = "";
                        j++;
                    }
                }
                var rFirstColumn = worksheet.Range("A6:A" + j).Merge();
                var rQuantity = worksheet.Range("H6:H" + j);
                worksheet.Cell("A6").Value = headersAndFooters["Materials"];

                rFirstColumn.Style.Fill.BackgroundColor = XLColor.Aqua;
                rFirstColumn.Style.Alignment.TextRotation = 180;
                rFirstColumn.Style.Font.Bold = true;
                rFirstColumn.Style.Font.Italic = true;
                rFirstColumn.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                rFirstColumn.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                rQuantity.Style.Fill.BackgroundColor = XLColor.Aqua;
                j += 1;
                var rangeMerged101 = worksheet.Range("A" + j + ":E" + j).Merge();
                worksheet.Cell("A" + j).Value = headersAndFooters["Specs"];
                rangeMerged101.Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell("A" + j).Style.Font.Bold = true;
                worksheet.Cell("A" + j).Style.Font.Italic = true;

                var rangeMerged102 = worksheet.Range("F" + j + ":I" + j).Merge();
                worksheet.Cell("F" + j).Value = headersAndFooters["Suppliers"];
                rangeMerged102.Style.Fill.BackgroundColor = XLColor.Aqua;
                worksheet.Cell("F" + j).Style.Font.Bold = true;
                worksheet.Cell("F" + j).Style.Font.Italic = true;
                worksheet.Cell("F" + j).Style.Font.Bold = true;
                worksheet.Cell("F" + j).Style.Font.Italic = true;
                j += 1;
                i = 1;
                worksheet.Cell("A" + j).Value = i;
                worksheet.Cell("A" + j).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell("F" + j).Value = i;
                worksheet.Cell("F" + j).Style.Fill.BackgroundColor = XLColor.Aqua;
                var rangeMerged103 = worksheet.Range("B" + j + ":D" + j).Merge();
                var rangeMerged103a = worksheet.Range("G" + j + ":I" + j).Merge();
                j += 1;
                i++;
                worksheet.Cell("A" + j).Value = i;
                worksheet.Cell("A" + j).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell("F" + j).Value = i;
                worksheet.Cell("F" + j).Style.Fill.BackgroundColor = XLColor.Aqua;
                var rangeMerged103b = worksheet.Range("B" + j + ":D" + j).Merge();
                var rangeMerged103c = worksheet.Range("G" + j + ":I" + j).Merge();
                j += 1;
                i++;
                worksheet.Cell("A" + j).Value = i;
                worksheet.Cell("A" + j).Style.Fill.BackgroundColor = XLColor.LightYellow;
                worksheet.Cell("F" + j).Value = i;
                worksheet.Cell("F" + j).Style.Fill.BackgroundColor = XLColor.Aqua;
                var rangeMerged103d = worksheet.Range("B" + j + ":D" + j).Merge();
                var rangeMerged103e = worksheet.Range("G" + j + ":I" + j).Merge();
                j += 1;
                var l = j;
                l += 4;
                var rangeMerged104 = worksheet.Range("A" + j + ":D" + j).Merge();
                worksheet.Cell("A" + j).Value = headersAndFooters["Remarks"];
                worksheet.Cell("A" + j).Style.Fill.BackgroundColor = XLColor.Aqua;
                worksheet.Cell("A" + j).Style.Font.Bold = true;
                worksheet.Cell("A" + j).Style.Font.Italic = true;
                var rangeMerged104b = worksheet.Range("E" + j + ":E" + l).Merge();
                var rangeMerged104c = worksheet.Range("F" + j + ":G" + l).Merge();
                var rangeMerged104d = worksheet.Range("H" + j + ":H" + l).Merge();
                var rangeMerged104e = worksheet.Range("I" + j + ":I" + l).Merge();
                i = 0;
                j += 1;
                i++;

                worksheet.Cell("A" + j).Value = i;
                worksheet.Cell("A" + j).Style.Fill.BackgroundColor = XLColor.Aqua;
                var rangeMerged104a = worksheet.Range("B" + j + ":D" + j).Merge();

                j += 1;
                i++;
                worksheet.Cell("A" + j).Value = i;
                worksheet.Cell("A" + j).Style.Fill.BackgroundColor = XLColor.Aqua;
                var rangeMerged104f = worksheet.Range("B" + j + ":D" + j).Merge();
                j += 1;
                i++;
                worksheet.Cell("A" + j).Value = i;
                worksheet.Cell("A" + j).Style.Fill.BackgroundColor = XLColor.Aqua;
                var rangeMerged104g = worksheet.Range("B" + j + ":D" + j).Merge();
                j += 1;
                i++;
                worksheet.Cell("A" + j).Value = i;
                worksheet.Cell("A" + j).Style.Fill.BackgroundColor = XLColor.Aqua;
                var rangeMerged104h = worksheet.Range("B" + j + ":D" + j).Merge();
                j += 1;
                i++;
                worksheet.Cell("A" + j).Value = i;
                worksheet.Cell("A" + j).Style.Fill.BackgroundColor = XLColor.Aqua;
                var rangeMerged105 = worksheet.Range("B" + j + ":D" + j).Merge();
                var rangeMerged105a = worksheet.Range("E" + j + ":E" + j).Merge();
                var rangeMerged105b = worksheet.Range("F" + j + ":G" + j).Merge();
                var rangeMerged105c = worksheet.Range("H" + j + ":H" + j).Merge();
                var rangeMerged105d = worksheet.Range("I" + j + ":I" + j).Merge();
                worksheet.Cell("E" + j).Value = headersAndFooters["Date"];
                worksheet.Cell("E" + j).Style.Font.Bold = true;
                worksheet.Cell("E" + j).Style.Font.Italic = true;
                worksheet.Cell("E" + j).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("E" + j).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("F" + j).Value = headersAndFooters["PM"];
                worksheet.Cell("F" + j).Style.Font.Bold = true;
                worksheet.Cell("F" + j).Style.Font.Italic = true;
                worksheet.Cell("F" + j).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("F" + j).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("H" + j).Value = headersAndFooters["Procurement"];
                worksheet.Cell("H" + j).Style.Font.Bold = true;
                worksheet.Cell("H" + j).Style.Font.Italic = true;
                worksheet.Cell("H" + j).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("H" + j).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("I" + j).Value = headersAndFooters["PM1"];
                worksheet.Cell("I" + j).Style.Font.Bold = true;
                worksheet.Cell("I" + j).Style.Font.Italic = true;
                worksheet.Cell("I" + j).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell("I" + j).Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                j += 1;
                var rangeLastRow = worksheet.Range("A" + j + ":I" + j).Merge();
                rangeLastRow.Style.Fill.BackgroundColor = XLColor.LightYellow1;
                var rangeRows = worksheet.Range("A1" + ":I" + j);

                rangeRows.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                rangeRows.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                rangeRows.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                rangeRows.Style.Border.RightBorder = XLBorderStyleValues.Thin;


                worksheet.Columns(1, 10).AdjustToContents();
                //worksheet.Column(1).Width = 20;
                string filePath = Convert.ToString(headersAndFooters["ReportGeneratedPath"]) +
                    "Indent_" + DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() +
                    DateTime.Now.Day.ToString() + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() +
                        DateTime.Now.Second.ToString() + ".xlsx";
                filePathLocation = filePath;
                workBook.SaveAs(filePath);
                this.Cursor = null;
                MessageBox.Show("The Indent file has been created.", "Order Management System",
                    MessageBoxButton.OK, MessageBoxImage.Information);

                ////log.Info("Indent has been generated..");
            }
            catch (Exception ex)
            {
                //Console.WriteLine("An Exception occurred. Kindly contact the Administrator");
                //////log.Error("Error while generating for Bot Report");
                //////log.Error(ex.Message);
                MessageBox.Show("An error occurred. Please contact the administrator", "Order Management System",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                ////log.Error("Error while gererating indent report: " + ex.StackTrace);
            }
        }
       
        protected override void OnClosing(CancelEventArgs e)
        {
            Menu menu = new Menu(_login);
            menu.Show();
        }

        private void btn_upload_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string filePath = string.Empty;
                OpenFileDialog openFileDialog = new OpenFileDialog();
                if (openFileDialog.ShowDialog() == true)
                    filePath = openFileDialog.FileName;
                //File.ReadAllText(openFileDialog.FileName);
                List<ExcelIndent> lstInputData = new List<ExcelIndent>();
                //DirectoryInfo dInfo = new DirectoryInfo(path);
                //FileInfo[] files = dInfo.GetFiles("*.xlsx");
                //if (files.Length == 0)
                //    MessageBox.Show("There are no files to process. Kindly move the files in the location.",
                //      "Order Management System", MessageBoxButton.OK, MessageBoxImage.Information);
                //foreach (FileInfo file in files)
                //{

                //    string filePath = path + "\\" + file.Name;
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    PullIndentData(lstInputData, stream);
                }
                File.SetAttributes(filePath, FileAttributes.Normal);
                string targetPath = Convert.ToString(ConfigurationManager.AppSettings["TargetReportPath"]);
                ////var fileName = file.Name.Split('.');
                ////targetPath = targetPath + "\\" + fileName[0] + DateTime.Now.ToString()+ "."+fileName[1];
                //targetPath = targetPath + "\\" + openFileDialog.SafeFileName;
                //File.Move(filePath, targetPath, true);
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
        private void PullIndentData(List<ExcelIndent> lstIndent, FileStream stream)
        {
            //ExcelIndent loadExcelIndent = null;
            List<ExcelIndent> lstGridIndent = new List<ExcelIndent>();
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
                                    if (!Convert.ToString(reader.GetValue(column)).Contains("Item Category Name"))
                                    {

                                        lst.Add(Convert.ToString(reader.GetValue(column)));
                                        isHeader = false;
                                    }
                                    //Console.WriteLine(reader.GetValue(column));//Get Value returns object
                                    else
                                    {
                                        isHeader = true;
                                        break;
                                    }
                                //else
                                //{
                                //    isHeader = true;
                                //    break;
                                //}


                            }
                            if (!isHeader)
                                if (lst.Count > 0)
                                    _readData.Add(lst);

                        }
                        //foreach (string s in lst)
                        //{

                        foreach (List<string> lst in _readData)
                        {

                            //  decimal totalAmount = 0;
                            if (lst.Count > 0)
                            {
                                //if (lst[4] == "1")
                                //{
                                //if (loadExcelIndent != null)
                                //{
                                //    // CreateIndent(loadExcelIndent);
                                //}
                                //loadExcelIndent = new LoadExcelIndent();
                                ////lst[1] = lst[1].Contains('-') ? lst[1].Replace('-', '/') : lst[1];
                                ////var dateComponent = lst[1].Split(' ')[0].Split('/');
                                //lst[1] = dateComponent[1] + "/" + dateComponent[0] + "/" + dateComponent[2];
                                //var raisedBy = (from i in orderManagementContext.UserMaster
                                //                where i.Email == lst[0]
                                //                select i).FirstOrDefault();
                                //loadExcelIndent.RaisedBy = raisedBy.EmployeeId;
                                //loadExcelIndent.Date = Convert.ToDateTime(lst[1]);
                                //loadExcelIndent.Location = lst[2];
                                //loadExcelIndent.ApproverName = lst[3];
                                //var approver = (from i in orderManagementContext.Employee
                                //                where i.FirstName.Contains(lst[3])
                                //                select i).FirstOrDefault();
                                //loadExcelIndent.ApprovalID = approver.EmployeeId;
                                //loadExcelIndent.CreateDate = DateTime.Now;

                                //lstGridIndent = new List<ExcelIndent>();
                                //ExcelIndent _indent = new ExcelIndent();
                                //_indent.SlNo = Convert.ToInt32(lst[4]);
                                //_indent.ItemCategoryName = lst[5];
                                //_indent.ItemCode = lst[6];
                                //_indent.ItemName = lst[7];
                                //_indent.Description = lst[8];
                                //_indent.Technical_Specifications = lst[9];
                                //_indent.Units = lst[10];
                                //_indent.Quantity = Convert.ToInt32(lst[11]);
                                //_indent.Remarks = lst[12];

                                //    lstGridIndent.Add(_indent);
                                //    loadExcelIndent.ExcelIndents = lstGridIndent;
                                //}
                                //else
                                //{
                                // lstGridIndent = new List<ExcelIndent>();
                                ExcelIndent _indent = new ExcelIndent();
                                //_indent.SlNo = Convert.ToInt32(lst[4]);
                                _indent.ItemCategoryName = lst[0];
                                _indent.ItemCategoryDescription = lst[1];
                                _indent.ItemCode = lst[2];
                                _indent.ItemName = lst[3];
                                _indent.Description = lst[4];
                                _indent.Technical_Specifications = lst[5];
                                _indent.Units = lst[6];
                                _indent.UnitsDescription = lst[7];
                                _indent.Quantity = Convert.ToInt32(lst[8]);
                                _indent.CreatedBy = lst[9];
                                if (lst.Count > 10)
                                {
                                    _indent.Remarks = lst[10];
                                }
                                else
                                {
                                    _indent.Remarks = "";
                                }
                                // raisedBy = lst[9];
                                lstGridIndent.Add(_indent);
                                // loadExcelIndent.ExcelIndents = lstGridIndent;
                            }
                        }
                        //}
                    } while (reader.NextResult()); //Move to NEXT SHEET

                    if (lstGridIndent != null)
                    {
                        // txt_raised_by.Text = raisedBy;
                        bool entryMade = CreateEntry(lstGridIndent);

                        if (entryMade)
                        {
                            foreach (var excelIndent in lstGridIndent)
                            {
                                GridIndent gridIndent = new GridIndent();
                                gridIndent.SlNo = gridIndents.Count + 1;
                                gridIndent.ItemCategoryName = excelIndent.ItemCategoryName;
                                gridIndent.ItemCode = excelIndent.ItemCode;
                                gridIndent.ItemName = excelIndent.ItemName;
                                gridIndent.Quantity = excelIndent.Quantity;
                                gridIndent.Remarks = excelIndent.Remarks;
                                gridIndent.Units = excelIndent.Units;
                                gridIndent.Description = excelIndent.Description;

                                bool itemPresent = false;
                                foreach (var i in gridIndents)
                                {
                                    if (i.ItemName == gridIndent.ItemName && i.ItemCode == gridIndent.ItemCode && i.Description == gridIndent.Description && i.Technical_Specifications == gridIndent.Technical_Specifications && itemPresent == false)
                                    {
                                        itemPresent = true;
                                        i.Quantity = i.Quantity + gridIndent.Quantity;
                                    }
                                }
                                if (!itemPresent)
                                {
                                    gridIndents.Add(gridIndent);
                                }
                            }
                            grid_indentdata.ItemsSource = null;
                            grid_indentdata.ItemsSource = gridIndents;
                            gridSelectedIndex = -1;
                            // CreateIndent(saveIndent);
                        }
                    }
                }
            }
            //return null;// inputData;
        }

        private bool CreateEntry(List<ExcelIndent> excelIndents)
        {

            bool dataUpdated = false;
            try
            {
                foreach (var excelIndent in excelIndents)
                {
                    var isItemCategoryPresent = (from i in orderManagementContext.ItemCategory
                                                 where i.ItemCategoryName == excelIndent.ItemCategoryName
                                                 select i.ItemCategoryId).FirstOrDefault();
                    if (isItemCategoryPresent == 0)
                    {
                        orderManagementContext.ItemCategory.Add(
                       new ItemCategory
                       {
                           ItemCategoryName = excelIndent.ItemCategoryName,
                           CreatedDate = DateTime.Now,
                           CreatedBy = _login.EmployeeID,
                           Description = excelIndent.ItemCategoryDescription
                       }
                       );
                        orderManagementContext.SaveChanges();
                    }

                    var isItemCodePresent = (from i in orderManagementContext.ItemMaster
                                             from ic in orderManagementContext.ItemCategory
                                             where i.ItemCategoryId == ic.ItemCategoryId && ic.ItemCategoryName == excelIndent.ItemCategoryName && i.ItemCode == excelIndent.ItemCode
                                             select i.ItemMasterId).FirstOrDefault();

                    if (isItemCodePresent == 0)
                    {
                        var itemCategoryId = (from i in orderManagementContext.ItemCategory
                                              where i.ItemCategoryName == excelIndent.ItemCategoryName
                                              select i.ItemCategoryId).FirstOrDefault();

                        orderManagementContext.ItemMaster.Add(
                            new ItemMaster
                            {
                                ItemCode = excelIndent.ItemCode,
                                ItemName = excelIndent.ItemName,
                                Description = excelIndent.Description,
                                TechnicalSpecification = excelIndent.Technical_Specifications,
                                ItemCategoryId = itemCategoryId,
                                CreatedDate = DateTime.Now,
                                CreatedBy = _login.EmployeeID
                            });
                        orderManagementContext.SaveChanges();
                    }


                    var isUnitPresent = (from u in orderManagementContext.UnitMaster
                                         where u.Unit == excelIndent.Units
                                         select u.UnitMasterId).FirstOrDefault();
                    if (isUnitPresent == 0)
                    {
                        orderManagementContext.UnitMaster.Add(
                        new UnitMaster
                        {
                            Unit = excelIndent.Units,
                            Description = excelIndent.UnitsDescription,
                            CreatedBy = _login.EmployeeID,
                            CreatedDate = DateTime.Now
                        });
                        orderManagementContext.SaveChanges();
                    }
                    dataUpdated = true;
                }
                dataUpdated = true;
                return dataUpdated;
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured during item details data upload  " + ex.Message, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                //log.Error("Error while creating approval via upload : " + ex.StackTrace);
                dataUpdated = false;
                return dataUpdated;
            }
        }

        private void CreateIndent(SaveIndent saveIndent)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlConnection"].ToString()))
                {
                    connection.Open();
                    SqlCommand testCMD = new SqlCommand("create_indent", connection);
                    testCMD.CommandType = CommandType.StoredProcedure;

                    testCMD.Parameters.Add(new SqlParameter("@Date", System.Data.SqlDbType.VarChar, 50) { Value = saveIndent.Date });
                    testCMD.Parameters.Add(new SqlParameter("@Location", System.Data.SqlDbType.VarChar, 50) { Value = saveIndent.LocationCode });
                    testCMD.Parameters.Add(new SqlParameter("@RaisedBy", System.Data.SqlDbType.BigInt, 50) { Value = saveIndent.RaisedBy });
                    testCMD.Parameters.Add(new SqlParameter("@CreatedDate", System.Data.SqlDbType.DateTime, 50) { Value = saveIndent.CreateDate });
                    testCMD.Parameters.Add(new SqlParameter("@IndentId", System.Data.SqlDbType.BigInt, 50) { Value = saveIndent.IndentId });
                    testCMD.Parameters["@IndentId"].Direction = ParameterDirection.Output;

                    testCMD.ExecuteNonQuery(); // read output value from @NewId 
                    saveIndent.IndentId = Convert.ToInt32(testCMD.Parameters["@IndentId"].Value);

                    indentNo = Convert.ToInt32(testCMD.Parameters["@IndentId"].Value);

                    foreach (var i in saveIndent.GridIndents)
                    {
                        string createDate = saveIndent.CreateDate.Month + "/" + saveIndent.CreateDate.Day + "/" + saveIndent.CreateDate.Year;
                        SqlCommand testCMD1 = new SqlCommand("create_indentDetails", connection);
                        testCMD1.CommandType = CommandType.StoredProcedure;
                        testCMD1.Parameters.Add(new SqlParameter("@IndentID", System.Data.SqlDbType.BigInt, 50) { Value = saveIndent.IndentId });
                        testCMD1.Parameters.Add(new SqlParameter("@ItemName", System.Data.SqlDbType.VarChar, 50) { Value = i.ItemCategoryName });
                        testCMD1.Parameters.Add(new SqlParameter("@ItemCode", System.Data.SqlDbType.VarChar, 50) { Value = i.ItemCode });
                        testCMD1.Parameters.Add(new SqlParameter("@Unit", System.Data.SqlDbType.VarChar, 50) { Value = i.Units });
                        testCMD1.Parameters.Add(new SqlParameter("@Quantity", System.Data.SqlDbType.VarChar, 50) { Value = i.Quantity });
                        testCMD1.Parameters.Add(new SqlParameter("@CreatedDate", System.Data.SqlDbType.DateTime, 50) { Value = createDate });
                        testCMD1.Parameters.Add(new SqlParameter("@CreatedBy", System.Data.SqlDbType.BigInt, 50) { Value = saveIndent.RaisedBy });
                        testCMD1.Parameters.Add(new SqlParameter("@Remarks", System.Data.SqlDbType.VarChar, 50) { Value = i.Remarks });
                        testCMD1.ExecuteNonQuery(); // read output value from @NewId 
                    }

                    SqlCommand testCMD2 = new SqlCommand("create_indent_Approval", connection);
                    testCMD2.CommandType = CommandType.StoredProcedure;
                    testCMD2.Parameters.Add(new SqlParameter("@IndentID", System.Data.SqlDbType.BigInt, 50) { Value = saveIndent.IndentId });
                    testCMD2.Parameters.Add(new SqlParameter("@ApprovalID", System.Data.SqlDbType.BigInt, 50) { Value = saveIndent.ApprovalID });
                    testCMD2.Parameters.Add(new SqlParameter("@ApprovalStatusID", System.Data.SqlDbType.BigInt, 50) { Value = 1 });
                    testCMD2.Parameters.Add(new SqlParameter("@CreatedDate", System.Data.SqlDbType.DateTime, 50) { Value = saveIndent.CreateDate });
                    testCMD2.Parameters.Add(new SqlParameter("@CreatedBy", System.Data.SqlDbType.BigInt, 50) { Value = saveIndent.RaisedBy });
                    testCMD2.ExecuteNonQuery();
                }
                //log.Info("Indent created and generated.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured during upload. " + ex.Message, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                //log.Error("Error while creating approval via upload : " + ex.StackTrace);
            }
        }
        private void btn_remove_indent_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (gridSelectedIndex >= 0)
                {
                    gridIndents.RemoveAt(gridSelectedIndex);
                    int count = 1;
                    foreach (var i in gridIndents)
                    {
                        i.SlNo = count;
                        count++;
                    }
                    grid_indentdata.ItemsSource = null;
                    grid_indentdata.ItemsSource = gridIndents;
                    gridSelectedIndex = -1;
                    ClearFields();
                }
                else
                {
                    MessageBox.Show("Please select a item");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred :" + ex.Message,
                                 "Order Management System",
                                     MessageBoxButton.OK,
                                         MessageBoxImage.Error);
                //log.Error("Error while removing Indent : " + ex.StackTrace);
            }
        }

        private void btn_add_category_Click(object sender, RoutedEventArgs e)
        {
            AddItemCategory addCategory = new AddItemCategory(_login);
            addCategory.Show();
        }

        private void btn_item_code_Click(object sender, RoutedEventArgs e)
        {
            AddItemCode addCode = new AddItemCode(_login);
            addCode.Show();
        }
        private void btn_add_unit_Click(object sender, RoutedEventArgs e)
        {
            AddUnit addUnit = new AddUnit(_login);
            addUnit.Show();
        }
        private void cbx_location_id_DropDownOpened(object sender, EventArgs e)
        {
            try
            {
                LoadLocationId();
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred during Location Load. " + ex.Message,
                                   "Order Management System",
                                       MessageBoxButton.OK,
                                           MessageBoxImage.Error);
                ////log.Error("Error while loading approval : " + ex.StackTrace);
            }
        }
    }
}