using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using OrderManagementTool.Models;
using OrderManagementTool.Models.Excel;
using OrderManagementTool.Models.Indent;
using OrderManagementTool.Models.LogIn;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ILog = log4net.ILog;
using LogManager = log4net.LogManager;

namespace OrderManagementTool
{
    /// <summary>
    /// Interaction logic for Indent.xaml
    /// </summary>
    public partial class Indent : Window
    {
        ILog log = LogManager.GetLogger(typeof(MainWindow));
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
        public Indent(Login login)
        {
            try
            {
                _login = login;
                log.Info("In Indent Screen...");
                InitializeComponent();
                LoadItemCategoryName();
                LoadApprovalStatus();
                txt_raised_by.Text = _login.UserEmail;
            }
            catch(Exception ex)
            {

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

        public Indent(Login login, long indentNo)
        {
            try
            {
                _login = login;
                log.Info("In Indent Screen...");
                InitializeComponent();
                LoadItemCategoryName();
                LoadApprovalStatus();
                txt_raised_by.Text = _login.UserEmail;
                GetIndent(indentNo);
            }
            catch(Exception ex)
            {

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

        private void GetIndent(long indentNo)
        {
            try
            {
                log.Info("Getting Indent infomration for Indent No: " + indentNo);
                using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlConnection"].ToString()))
                {
                    connection.Open();
                    SqlCommand testCMD = new SqlCommand("GetIndent", connection);
                    testCMD.CommandType = CommandType.StoredProcedure;

                    testCMD.Parameters.Add(new SqlParameter("@IndentID", System.Data.SqlDbType.BigInt, 50) { Value = indentNo });

                    // SqlDataReader dataReader = testCMD.ExecuteReader();
                    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(testCMD);

                    DataSet dataSet = new DataSet();
                    sqlDataAdapter.Fill(dataSet);

                    SaveIndent saveIndent = new SaveIndent();

                    int counter = 0;

                    while (counter < dataSet.Tables[0].Rows.Count)
                    {
                        saveIndent.Date = Convert.ToDateTime(dataSet.Tables[0].Rows[counter]["Date"]);
                        saveIndent.Location = Convert.ToString(dataSet.Tables[0].Rows[counter]["Location"]);
                        saveIndent.IndentRemarks = Convert.ToString(dataSet.Tables[0].Rows[counter]["Remarks"]);
                        saveIndent.ApproverName = Convert.ToString(dataSet.Tables[0].Rows[counter]["Approver"]);

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
                    txt_location.Text = saveIndent.Location;
                    cbx_approval_id.SelectedValue = saveIndent.ApproverName;

                    grid_indentdata.ItemsSource = null;
                    grid_indentdata.ItemsSource = gridIndents;

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured during save. " + ex.Message, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                log.Error("Error while fetching Indent information: " + ex.StackTrace);
            }
        }

        private void DataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //DataRowView rowview = grid_indentdata.SelectedItem as DataRowView;
            try
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
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred :" + ex.Message,
                                   "Order Management System",
                                       MessageBoxButton.OK,
                                           MessageBoxImage.Error);
                log.Error("Error while selection of items in datagrid: " + ex.StackTrace);
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
                log.Error("Error while selecting item category: " + ex.StackTrace);
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
                log.Error("Error while selecting item code: " + ex.StackTrace);
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
                log.Error("Error while selecting units: " + ex.StackTrace);
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
                log.Error("Error while selecting Item code: " + ex.StackTrace);
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                log.Info("Adding Indent information");
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
                log.Error("Indent infomration added");
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred :" + ex.Message,
                                   "Order Management System",
                                       MessageBoxButton.OK,
                                           MessageBoxImage.Error);
                log.Error("Error while adding Indent information: " + ex.StackTrace);
            }
        }

        private void LoadApprovalStatus()
        {
            try
            {
                log.Error("Loading Approval infomration");
                cbx_approval_id.Items.Clear();
                var exceptionList = new List<string> { "Clerk", "Supervisor" };
                var data = (from emp in orderManagementContext.Employee
                            join des in orderManagementContext.Designation
                            on emp.DesignationId equals des.DesignationId
                            select
                            new
                            {
                                emp.FirstName,
                                emp.LastName,
                                des.Designation1
                            })
                                 .Distinct().ToList();
                foreach (string s in exceptionList)
                {
                    var remove = data.Select(e => e.Designation1 == s).ToList();
                    int index = remove.IndexOf(true);
                    if (index >= 0)
                        data.RemoveAt(index);
                }

                foreach (var i in data)
                {
                    cbx_approval_id.Items.Add(i.FirstName.Trim() + " " + i.LastName.Trim());
                }
                log.Error("Approval Information loaded.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured during Approval ID fetch " + ex.Message, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                log.Error("Error while loading approval : " + ex.StackTrace);
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

        private void LoadApprovalId()
        {
            LoadApprovalStatus();
        }

        private void btn_clear_fields_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                log.Info("Clearing Fields");
                ClearFields();
                log.Info("Cleared fields");
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred :" + ex.Message,
                                   "Order Management System",
                                       MessageBoxButton.OK,
                                           MessageBoxImage.Error);
                log.Error("Error while clearing Indent information: " + ex.StackTrace);
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            log.Info("Adding data to grid");
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
                log.Info("Data added to grid");
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
                LoadApprovalId();
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred :" + ex.Message,
                                   "Order Management System",
                                       MessageBoxButton.OK,
                                           MessageBoxImage.Error);
                log.Error("Error while loading approval : " + ex.StackTrace);
            }
        }

        private void btn_create_indent_Click(object sender, RoutedEventArgs e)
        {
            this.Cursor = Cursors.Wait;
            log.Info("Creating Indent...");
            if (datepicker_date1.SelectedDate.ToString() == "")
            {
                MessageBox.Show("Please enter Valid Date");
                return;
            }
            if (cbx_approval_id.SelectedValue == null)
            {
                MessageBox.Show("Please select Approver");
                return;
            }
            if (txt_location.Text == "")
            {
                MessageBox.Show("Please Enter Location");
                return;
            }

            if (gridIndents.Count == 0)
            {
                MessageBox.Show("Please Enter Indent Details");
                return;
            }

            else
            {
                SaveIndent saveIndent = new SaveIndent();
                saveIndent.Date = datepicker_date1.SelectedDate;
                saveIndent.Location = txt_location.Text;
                saveIndent.RaisedBy = _login.EmployeeID;
                saveIndent.CreateDate = DateTime.Now;
                var approvalID = (from a in orderManagementContext.UserMaster
                                  where a.Email == cbx_approval_id.SelectedValue.ToString()
                                  select a.UserId).FirstOrDefault();
                saveIndent.ApprovalID = approvalID;
                // saveIndent.ApprovalID = 1;
                saveIndent.GridIndents = gridIndents;
                try
                {
                    using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlConnection"].ToString()))
                    {
                        connection.Open();
                        SqlCommand testCMD = new SqlCommand("create_indent", connection);
                        testCMD.CommandType = CommandType.StoredProcedure;

                        testCMD.Parameters.Add(new SqlParameter("@Date", System.Data.SqlDbType.VarChar, 50) { Value = saveIndent.Date });
                        testCMD.Parameters.Add(new SqlParameter("@Location", System.Data.SqlDbType.VarChar, 50) { Value = saveIndent.Location });
                        testCMD.Parameters.Add(new SqlParameter("@RaisedBy", System.Data.SqlDbType.BigInt, 50) { Value = saveIndent.RaisedBy });
                        testCMD.Parameters.Add(new SqlParameter("@CreatedDate", System.Data.SqlDbType.DateTime, 50) { Value = saveIndent.CreateDate });
                        testCMD.Parameters.Add(new SqlParameter("@IndentId", System.Data.SqlDbType.BigInt, 50) { Value = saveIndent.IndentId });
                        testCMD.Parameters["@IndentId"].Direction = ParameterDirection.Output;

                        testCMD.ExecuteNonQuery(); // read output value from @NewId 
                        saveIndent.IndentId = Convert.ToInt32(testCMD.Parameters["@IndentId"].Value);

                        indentNo= Convert.ToInt32(testCMD.Parameters["@IndentId"].Value);
                        indentNo= Convert.ToInt32(testCMD.Parameters["@IndentId"].Value);
                        foreach (var i in saveIndent.GridIndents)
                        {
                            SqlCommand testCMD1 = new SqlCommand("create_indentDetails", connection);
                            testCMD1.CommandType = CommandType.StoredProcedure;
                            testCMD1.Parameters.Add(new SqlParameter("@IndentID", System.Data.SqlDbType.BigInt, 50) { Value = saveIndent.IndentId });
                            testCMD1.Parameters.Add(new SqlParameter("@ItemName", System.Data.SqlDbType.VarChar, 50) { Value = i.ItemCategoryName });
                            testCMD1.Parameters.Add(new SqlParameter("@ItemCode", System.Data.SqlDbType.VarChar, 50) { Value = i.ItemCode });
                            testCMD1.Parameters.Add(new SqlParameter("@Unit", System.Data.SqlDbType.VarChar, 50) { Value = i.Units });
                            testCMD1.Parameters.Add(new SqlParameter("@Quantity", System.Data.SqlDbType.VarChar, 50) { Value = i.Quantity });
                            testCMD1.Parameters.Add(new SqlParameter("@CreatedDate", System.Data.SqlDbType.VarChar, 50) { Value = saveIndent.CreateDate });
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
                    GenerateIndent();
                    log.Info("Indent created and generated.");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("An error occured during save. " + ex.Message, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                    log.Error("Error while creating approval : " + ex.StackTrace);
                }
                finally
                {
                    this.Cursor = null;
                }
            }
        }

        private bool SendMail(string body)
        {
            log.Info("Sending Mail..");
            bool mailSent = false;
            string message = DateTime.Now + " In SendMail\n";

            using (MailMessage mm = new MailMessage())
            {
                mm.From = new MailAddress(Convert.ToString(ConfigurationManager.AppSettings["MailFrom"]));
                string[] _toAddress = Convert.ToString(ConfigurationManager.AppSettings["MailTo"]).Split(';');
                foreach (string address in _toAddress)
                {
                    mm.To.Add(address);
                }
                mm.Subject = ConfigurationManager.AppSettings["Subject"];
                body = body + ". Please click on the link to approve or deny the indent. " + ConfigurationManager.AppSettings["url"] + "/" + indentNo;
                mm.Body = body;
                mm.IsBodyHtml = false;

                mm.Attachments.Add(new System.Net.Mail.Attachment(filePathLocation));

                SmtpClient smtp = new SmtpClient();
                smtp.Host = ConfigurationManager.AppSettings["Host"];
                smtp.EnableSsl = true;
                NetworkCredential NetworkCred = new NetworkCredential(ConfigurationManager.AppSettings["Username"],
                    ConfigurationManager.AppSettings["Password"]);
                smtp.UseDefaultCredentials = true;
                smtp.Credentials = NetworkCred;
                smtp.Port = int.Parse(ConfigurationManager.AppSettings["Port"]);

                message = DateTime.Now + " Sending Mail\n";
                smtp.Send(mm);
                message = DateTime.Now + " Mail Sent\n";

                System.Threading.Thread.Sleep(3000);
                mailSent = true;
            }
            return mailSent;
            log.Info("Mail sent..");
        }

        private void btn_generate_report_Click(object sender, RoutedEventArgs e)
        {
            GenerateIndent();
        }
        private void GenerateIndent()
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

            GenerateIndent(_headers, headerData, gridIndents);
        }
        private Dictionary<string, string> GetHeaders()
        {
            log.Info("Reading Header information from AppConfig");
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

        private void GenerateIndent(Dictionary<string, string> headersAndFooters,
            IndentHeader headerData, List<GridIndent> gridIndents)
        {
            log.Info("Generating Indent Report..");
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
                MessageBox.Show("The Indent file has been created.", "Order Management System",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                SendMail("Indent has been generated");
                log.Info("Indent has been generated..");
            }
            catch (Exception ex)
            {
                //Console.WriteLine("An Exception occurred. Kindly contact the Administrator");
                //log.Error("Error while generating for Bot Report");
                //log.Error(ex.Message);
                MessageBox.Show("An error occurred. Please contact the administrator", "Order Management System",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                log.Error("Error while gererating indent report: " + ex.StackTrace);
            }
        }
        private static void GeneratePurchaseOrder(Dictionary<string, string> headersAndFooters, List<ExportIndent> gridIndents)
        {
            try
            {
                var workBook = new XLWorkbook();
                workBook.AddWorksheet("Report");
                var worksheet = workBook.Worksheet("Report");

                var imagePath = @"D:\Dinesh\Projects\GitHub\OrderManagementTool\Images\logo.png";
                var image = worksheet.AddPicture(imagePath)
                    .MoveTo(worksheet.Cell("A1"))
                    .Scale(.1).Placement = ClosedXML.Excel.Drawings.XLPicturePlacement.Move;
                var rangeMerged = worksheet.Range("A1:G1").Merge();
                rangeMerged.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                rangeMerged.Style.Border.TopBorder = XLBorderStyleValues.Medium;
                rangeMerged.Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                rangeMerged.Style.Border.RightBorder = XLBorderStyleValues.Medium;
                rangeMerged.Style.Font.Bold = true;
                rangeMerged.Style.Font.FontSize = 16;
                rangeMerged.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                rangeMerged.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                var row = worksheet.Row(1);
                row.Height = 40;
                var rangeMerged1 = worksheet.Range("A2:G2").Merge();
                rangeMerged1.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                rangeMerged1.Style.Border.TopBorder = XLBorderStyleValues.Medium;
                rangeMerged1.Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                rangeMerged1.Style.Border.RightBorder = XLBorderStyleValues.Medium;
                rangeMerged1.Style.Font.Bold = true;
                rangeMerged1.Style.Font.FontSize = 16;
                rangeMerged1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                rangeMerged1.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                var rangeMerged2 = worksheet.Range("A3:G3").Merge();
                rangeMerged2.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                rangeMerged2.Style.Border.TopBorder = XLBorderStyleValues.Medium;
                rangeMerged2.Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                rangeMerged2.Style.Border.RightBorder = XLBorderStyleValues.Medium;
                rangeMerged2.Style.Font.Bold = true;
                rangeMerged2.Style.Font.FontSize = 16;
                rangeMerged2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                rangeMerged2.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                worksheet.Cell("A1").Value = worksheet.Cell("A1").Value + headersAndFooters["CompanyHeader"];
                worksheet.Cell("A3").Value = headersAndFooters["Header"];
                worksheet.Cell("A4").Value = headersAndFooters["To"];
                worksheet.Cell("D4").Value = headersAndFooters["Info"];

                worksheet.Cell("B5").Value = "M/S GILIYAL INDUSTRIES";
                worksheet.Cell("B6").Value = "Address Line 1";
                worksheet.Cell("B7").Value = "Address Line 2";
                worksheet.Cell("B8").Value = "Address Line 3";
                worksheet.Cell("B9").Value = "Address Line 4";
                worksheet.Cell("F5").Value = headersAndFooters["PONo"];
                worksheet.Cell("F6").Value = headersAndFooters["PODate"];
                worksheet.Cell("F7").Value = headersAndFooters["RefNo"];
                worksheet.Cell("F8").Value = headersAndFooters["RefDate"];
                worksheet.Cell("F9").Value = headersAndFooters["Attn"];
                worksheet.Cell("G5").Value = "PO Number";
                worksheet.Cell("G6").Value = "PO Date";
                worksheet.Cell("G7").Value = "Ref No";
                worksheet.Cell("G8").Value = "Ref Date";
                worksheet.Cell("G9").Value = "Attention";

                var rangeMerged3 = worksheet.Range("A10:C10").Merge();
                worksheet.Cell("A10").Value = headersAndFooters["Materials"];

                worksheet.Cell("B11").Value = headersAndFooters["CompanyName"];
                worksheet.Cell("B12").Value = headersAndFooters["CompanyAddressLine1"];
                worksheet.Cell("B13").Value = headersAndFooters["CompanyAddressLine2"];
                worksheet.Cell("B14").Value = headersAndFooters["CompanyContactNo"];
                worksheet.Cell("F10").Value = headersAndFooters["Remarks"];
                worksheet.Cell("F11").Value = headersAndFooters["GSTNo"];
                worksheet.Cell("F12").Value = headersAndFooters["IECNo"];
                worksheet.Cell("F13").Value = headersAndFooters["Page"];
                // worksheet.Cell("F9").Value = headersAndFooters["Attn"];
                worksheet.Cell("G10").Value = "Remarks";
                worksheet.Cell("G11").Value = "GST #";
                worksheet.Cell("G12").Value = "IEC #";
                worksheet.Cell("G13").Value = "1 Of 1";
                worksheet.Cell("G14").Value = "Attention";

                worksheet.Cell("A16").Value = "S.No";
                worksheet.Cell("B16").Value = "Description";
                worksheet.Cell("C16").Value = "Qty";
                worksheet.Cell("D16").Value = "Units";
                worksheet.Cell("E16").Value = "Unit Price (INR)";
                worksheet.Cell("F16").Value = "HSN";
                worksheet.Cell("G16").Value = "Total Price (INR)";
                int j = 17;
                int i = 1;
                decimal total = 0;
                foreach (ExportIndent indent in gridIndents)
                {
                    worksheet.Cell("A" + j).Value = i;
                    worksheet.Cell("B" + j).Value = indent.Description;
                    worksheet.Cell("C" + j).Value = indent.Quantity;
                    worksheet.Cell("D" + j).Value = indent.Units;
                    worksheet.Cell("E" + j).Value = indent.UnitPrice;
                    worksheet.Cell("F" + j).Value = indent.HSN;
                    worksheet.Cell("G" + j).Value = indent.TotalPrice;
                    total += indent.TotalPrice;
                    j++;
                    i++;
                }

                worksheet.Cell("B" + j).Value = headersAndFooters["SubTotal"];
                worksheet.Cell("G" + j).Value = total;

                j += 1;
                decimal gst = Math.Round(total * Convert.ToDecimal(.18));
                decimal finalTotal = total + gst;
                worksheet.Cell("B" + j).Value = headersAndFooters["GST"];
                worksheet.Cell("G" + j).Value = Convert.ToString(gst);

                j += 1;
                worksheet.Cell("B" + j).Value = headersAndFooters["FinalTotal"];
                worksheet.Cell("G" + j).Value = Convert.ToString(finalTotal);

                j += 1;
                var rangeMerged4 = worksheet.Range("B" + j + ":G" + j).Merge();
                worksheet.Cell("B" + j).Value = headersAndFooters["Words"];

                j += 2;
                var rangeMerged101 = worksheet.Range("A" + j + ":G" + j).Merge();
                worksheet.Cell("A" + j).Value = headersAndFooters["Terms"];
                j += 1;
                var rangeMerged102 = worksheet.Range("A" + j + ":G" + j).Merge();
                worksheet.Cell("A" + j).Value = headersAndFooters["Terms1"];

                j += 1;
                var rangeMerged103 = worksheet.Range("A" + j + ":G" + j).Merge();
                worksheet.Cell("A" + j).Value = headersAndFooters["Terms2"];
                j += 1;
                var rangeMerged104 = worksheet.Range("A" + j + ":G" + j).Merge();
                worksheet.Cell("A" + j).Value = headersAndFooters["Terms3"];
                j += 1;
                var rangeMerged105 = worksheet.Range("A" + j + ":G" + j).Merge();
                worksheet.Cell("A" + j).Value = headersAndFooters["Terms4"];
                j += 1;
                var rangeMerged106 = worksheet.Range("A" + j + ":G" + j).Merge();
                worksheet.Cell("A" + j).Value = headersAndFooters["Terms5"];
                j += 3;
                var rangeMerged5 = worksheet.Range("A" + j + ":D" + j).Merge();
                worksheet.Cell("A" + j).Value = headersAndFooters["SpecialInstructions"];
                var rangeMerged6 = worksheet.Range("E" + j + ":G" + j).Merge();
                worksheet.Cell("E" + j).Value = headersAndFooters["Rejections"];
                j += 1;
                var rangeMerged7 = worksheet.Range("A" + j + ":D" + j).Merge();
                worksheet.Cell("A" + j).Value = headersAndFooters["SplLine1"];
                worksheet.Cell("A" + j).Style.Alignment.WrapText = true;
                var rangeMerged8 = worksheet.Range("E" + j + ":G" + j).Merge();
                worksheet.Cell("E" + j).Value = headersAndFooters["RejLine1"];
                worksheet.Cell("E" + j).Style.Alignment.WrapText = true;
                row = worksheet.Row(1);
                row.Height = 40;
                j += 1;
                var rangeMerged9 = worksheet.Range("A" + j + ":D" + j).Merge();
                worksheet.Cell("A" + j).Value = headersAndFooters["SplLine2"];
                worksheet.Cell("A" + j).Style.Alignment.WrapText = true;
                j += 3;
                worksheet.Cell("G" + j).Value = headersAndFooters["For"] + " " + headersAndFooters["CompanyHeader"];
                j += 3;
                worksheet.Cell("G" + j).Value = headersAndFooters["Sign"];

                worksheet.Columns(1, 10).AdjustToContents();
                //worksheet.Column(1).Width = 20;
                string filePath = Convert.ToString(headersAndFooters["ReportGeneratedPath"]) +
                    "Indent_" + DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() +
                    DateTime.Now.Day.ToString() + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() +
                        DateTime.Now.Second.ToString() + ".xlsx";

                workBook.SaveAs(filePath);
            }
            catch (Exception ex)
            {
                //Console.WriteLine("An Exception occurred. Kindly contact the Administrator");
                //log.Error("Error while generating for Bot Report");
                //log.Error(ex.Message);
            }
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            Menu menu = new Menu(_login);
            menu.Show();
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
                log.Error("Error while removing Indent : " + ex.StackTrace);
            }
        }
    }
}