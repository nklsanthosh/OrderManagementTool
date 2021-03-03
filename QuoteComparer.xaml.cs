using ClosedXML.Excel;
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
using System.Net;
using System.Net.Mail;
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
using ViewIndent = OrderManagementTool.Models.Purchase_Order.ViewIndent;

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
        //private int gridSelectedIndex;
        OrderManagementContext orderManagementContext = new OrderManagementContext();
        private List<GridIndent> gridIndents = new List<GridIndent>();
        // public BindableCollection<string> ItemName { get; set; }
        private readonly Login _login;
        private static string filePathLocation;
        public long indentNo;
        // public string mailFrom;
        public string mailTo;
        //private bool isGridReadOnly = false;
        List<Poitem> poitems_1 = new List<Poitem>();
        List<Poitem> poitems_2 = new List<Poitem>();
        List<Poitem> poitems_3 = new List<Poitem>();
        List<Poitem> poitems_4 = new List<Poitem>();
        private long poID;

        public QuoteComparer(Login login)
        {
            _login = login;
            InitializeComponent();
            txt_indent_no.IsReadOnly = false;
            LoadApprovalStatus();
            cbx_ApprovalStatus_id.SelectedValue = 1;
            cbx_ApprovalStatus_id.IsEnabled = false;
            FillIndent();
            DisableCheckBoxes();
        }

        public QuoteComparer(Login login, string IndentNo)
        {
            _login = login;
            InitializeComponent();
            txt_indent_no.IsReadOnly = false;
            txt_indent_no.Text = IndentNo.ToString();
            LoadApprovalStatus();
            cbx_ApprovalStatus_id.SelectedValue = 1;
            cbx_ApprovalStatus_id.IsEnabled = false;
            FillIndent();
            DisableCheckBoxes();
        }

        public QuoteComparer(Login login, long PO_ID)
        {
            _login = login;
            InitializeComponent();
            GetPurchaseOrder(PO_ID);
            FillIndent();
            FillApprovalStatusLevel(txt_PO_no.Text.Trim().ToString());
            btn_Create_PO.IsEnabled = false;
            DisableUploadButtons();
        }

        private void DisableCheckBoxes()
        {
            checkbox_Approve1.IsEnabled = false;
            checkbox_Approve2.IsEnabled = false;
            checkbox_Approve3.IsEnabled = false;
        }

        private void EnableCheckBoxes()
        {
            checkbox_Approve1.IsEnabled = true;
            checkbox_Approve2.IsEnabled = true;
            checkbox_Approve3.IsEnabled = true;
        }

        private void DisableUploadButtons()
        {
            btn_upload_1.IsEnabled = false;
            btn_upload_2.IsEnabled = false;
            btn_upload_3.IsEnabled = false;
            btn_quotation_upload_1.IsEnabled = false;
            btn_quotation_upload_2.IsEnabled = false;
            btn_quotation_upload_3.IsEnabled = false;
        }

        private void GetPurchaseOrder(long PO_ID)
        {
            try
            {
                PORetrival pORetrival = new PORetrival();
                ////log.Info("Getting Indent infomration for Indent No: " + indentNo);
                using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlConnection"].ToString()))
                {
                    connection.Open();

                    SqlCommand testCMD = new SqlCommand("GetPurchaseOrder", connection);
                    testCMD.CommandType = CommandType.StoredProcedure;

                    testCMD.Parameters.Add(new SqlParameter("@POId", System.Data.SqlDbType.BigInt, 50) { Value = PO_ID });

                    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(testCMD);

                    DataSet dataSet = new DataSet();
                    sqlDataAdapter.Fill(dataSet);

                    int counter = 0;

                    List<PoItemwithQuotation> poItemwithQuotation = new List<PoItemwithQuotation>();


                    while (counter < dataSet.Tables[0].Rows.Count)
                    {
                        pORetrival.PO_ID = Convert.ToInt64(dataSet.Tables[0].Rows[counter]["PO_ID"]);
                        pORetrival.Indent_No = Convert.ToInt64(dataSet.Tables[0].Rows[counter]["IndentID"]);
                        pORetrival.Approval_Status_Id = Convert.ToInt64(dataSet.Tables[0].Rows[counter]["ApprovalStatusId"]);
                        pORetrival.Approver_Id = Convert.ToInt64(dataSet.Tables[0].Rows[counter]["ApproverId"]);

                        PoItemwithQuotation p1 = new PoItemwithQuotation();
                        p1.Offer_Number = Convert.ToString(dataSet.Tables[0].Rows[counter]["Offer_Number"]);
                        p1.Description = Convert.ToString(dataSet.Tables[0].Rows[counter]["Description"]);
                        p1.Vendor_Name = Convert.ToString(dataSet.Tables[0].Rows[counter]["Vendor_Name"]);
                        p1.Vendor_Code = Convert.ToString(dataSet.Tables[0].Rows[counter]["Vendor_Code"]);
                        p1.Offer_Date = Convert.ToString(dataSet.Tables[0].Rows[counter]["Offer_Date"]);
                        p1.Contact_No = Convert.ToString(dataSet.Tables[0].Rows[counter]["Contact_Number"]);
                        p1.Contact_Person = Convert.ToString(dataSet.Tables[0].Rows[counter]["Contact_Person"]);
                        p1.GST_Value = Convert.ToDecimal(dataSet.Tables[0].Rows[counter]["GST_Value"]);
                        p1.HSN = Convert.ToInt32(dataSet.Tables[0].Rows[counter]["HSN"]);
                        p1.Quantity = Convert.ToInt32(dataSet.Tables[0].Rows[counter]["Quantity"]);
                        p1.Q_No = Convert.ToInt32(dataSet.Tables[0].Rows[counter]["Q_No"]);
                        p1.Units = Convert.ToString(dataSet.Tables[0].Rows[counter]["Units"]);
                        p1.Unit_Price = Convert.ToInt64(dataSet.Tables[0].Rows[counter]["Unit_Price"]);
                        p1.Total_Price = Convert.ToDecimal(dataSet.Tables[0].Rows[counter]["Total_Price"]);
                        poItemwithQuotation.Add(p1);
                        counter++;
                    }
                    dataSet.Dispose();

                    if (poItemwithQuotation.Count > 0)
                    {
                        foreach (var i in poItemwithQuotation)
                        {
                            if (i.Q_No == 1)
                            {
                                Poitem poitem = new Poitem();
                                poitem.Sl_NO = poitems_1.Count + 1;
                                poitem.Offer_Number = i.Offer_Number;
                                poitem.Vendor_Name = i.Vendor_Name;
                                poitem.Vendor_Code = i.Vendor_Code;
                                poitem.Offer_Date = i.Offer_Date;
                                poitem.Contact_No = i.Contact_No;
                                poitem.Contact_Person = i.Contact_Person;
                                poitem.Description = i.Description;
                                poitem.Quantity = i.Quantity;
                                poitem.Units = i.Units;
                                poitem.Total_Price = i.Total_Price;
                                poitem.Unit_Price = i.Unit_Price;
                                poitem.HSN = i.HSN;
                                poitem.GST_Value = i.GST_Value;
                                poitems_1.Add(poitem);
                            }
                            else if (i.Q_No == 2)
                            {
                                Poitem poitem = new Poitem();
                                poitem.Offer_Number = i.Offer_Number;
                                poitem.Vendor_Name = i.Vendor_Name;
                                poitem.Vendor_Code = i.Vendor_Code;
                                poitem.Offer_Date = i.Offer_Date;
                                poitem.Contact_No = i.Contact_No;
                                poitem.Contact_Person = i.Contact_Person;
                                poitem.Description = i.Description;
                                poitem.Quantity = i.Quantity;
                                poitem.Units = i.Units;
                                poitem.Total_Price = i.Total_Price;
                                poitem.Unit_Price = i.Unit_Price;
                                poitem.HSN = i.HSN;
                                poitem.GST_Value = i.GST_Value;
                                poitems_2.Add(poitem);
                            }
                            else if (i.Q_No == 3)
                            {
                                Poitem poitem = new Poitem();
                                poitem.Offer_Number = i.Offer_Number;
                                poitem.Vendor_Name = i.Vendor_Name;
                                poitem.Vendor_Code = i.Vendor_Code;
                                poitem.Offer_Date = i.Offer_Date;
                                poitem.Contact_No = i.Contact_No;
                                poitem.Contact_Person = i.Contact_Person;
                                poitem.Description = i.Description;
                                poitem.Quantity = i.Quantity;
                                poitem.Units = i.Units;
                                poitem.Total_Price = i.Total_Price;
                                poitem.Unit_Price = i.Unit_Price;
                                poitem.HSN = i.HSN;
                                poitem.GST_Value = i.GST_Value;
                                poitems_3.Add(poitem);
                            }
                            else if (i.Q_No == 4)
                            {
                                Poitem poitem = new Poitem();
                                poitem.Offer_Number = i.Offer_Number;
                                poitem.Vendor_Name = i.Vendor_Name;
                                poitem.Vendor_Code = i.Vendor_Code;
                                poitem.Offer_Date = i.Offer_Date;
                                poitem.Contact_No = i.Contact_No;
                                poitem.Contact_Person = i.Contact_Person;
                                poitem.Description = i.Description;
                                poitem.Quantity = i.Quantity;
                                poitem.Units = i.Units;
                                poitem.Total_Price = i.Total_Price;
                                poitem.Unit_Price = i.Unit_Price;
                                poitem.HSN = i.HSN;
                                poitem.GST_Value = i.GST_Value;
                                poitems_4.Add(poitem);
                                txt_GST.Text = i.GST_Value.ToString();
                            }
                        }
                        if (poitems_1.Count > 0)
                        {
                            grid_quote_1.ItemsSource = null;
                            grid_quote_1.ItemsSource = poitems_1;
                        }
                        if (poitems_2.Count > 0)
                        {
                            grid_quote_2.ItemsSource = null;
                            grid_quote_2.ItemsSource = poitems_2;
                        }
                        if (poitems_3.Count > 0)
                        {
                            grid_quote_3.ItemsSource = null;
                            grid_quote_3.ItemsSource = poitems_3;
                        }
                        if (poitems_4.Count > 0)
                        {
                            grid_po_confirmation.ItemsSource = null;
                            grid_po_confirmation.ItemsSource = poitems_4;
                        }

                        txt_PO_no.Text = pORetrival.PO_ID.ToString();
                        txt_indent_no.Text = pORetrival.Indent_No.ToString();
                        LoadApprovalStatus();
                        cbx_ApprovalStatus_id.SelectedValue = pORetrival.Approval_Status_Id;
                        if (pORetrival.Approval_Status_Id != 1)
                        {
                            DisableCheckBoxes();
                        }
                        else
                        {
                            EnableCheckBoxes();
                        }
                        btn_Approve.IsEnabled = false;
                        if (pORetrival.Approver_Id <= _login.EmployeeID)
                        {
                            if (pORetrival.Approver_Id == _login.EmployeeID)
                            {
                                if (pORetrival.Approval_Status_Id != 1)
                                {
                                    DisableFields();
                                    DisableCheckBoxes();
                                }
                                else
                                {
                                    return;
                                }
                            }
                            else
                            {
                                DisableFields();
                                DisableCheckBoxes();
                            }

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured during poData Retrival " + ex.Message, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadApprovalStatus()
        {
            try
            {
                cbx_ApprovalStatus_id.SelectedValuePath = "Key";
                cbx_ApprovalStatus_id.DisplayMemberPath = "Value";
                ////log.Error("Loading Approval infomration");
                cbx_ApprovalStatus_id.Items.Clear();

                var data = (from a in orderManagementContext.ApprovalStatus
                            select a).ToList();

                foreach (var i in data)
                {
                    cbx_ApprovalStatus_id.Items.Add(new KeyValuePair<long, string>(i.ApprovalStatusId, i.ApprovalStatus1));
                }
                ////log.Error("Approval Information loaded.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured during Approval Status fetch " + ex.Message, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                ////log.Error("Error while loading location code : " + ex.StackTrace);
            }
        }
        private void grid_po_confirmation_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
        private void FillIndent()
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlConnection"].ToString()))
                {
                    connection.Open();
                    SqlCommand testCMD = new SqlCommand("GetIndentByIndentNo", connection);
                    testCMD.CommandType = CommandType.StoredProcedure;

                    //testCMD.Parameters.Add(new SqlParameter("@UserId", System.Data.SqlDbType.BigInt, 50) { Value = _login.EmployeeID });
                    testCMD.Parameters.Add(new SqlParameter("@IndentID", System.Data.SqlDbType.VarChar, 300) { Value = txt_indent_no.Text.Trim().ToString() });

                    // SqlDataReader dataReader = testCMD.ExecuteReader();
                    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(testCMD);

                    DataSet dataSet = new DataSet();
                    sqlDataAdapter.Fill(dataSet);

                    int counter = 0;

                    if (viewIndents.Count > 0)
                    {
                        viewIndents.Clear();
                    }

                    while (counter < dataSet.Tables[0].Rows.Count)
                    {
                        ViewIndent viewIndent = new ViewIndent();
                        viewIndent.Sl_No = counter + 1;
                        // viewIndent.IndentId = (long)Convert.ToInt64(dataSet.Tables[0].Rows[counter]["IndentID"]);
                        viewIndent.ApproverName = Convert.ToString(dataSet.Tables[0].Rows[counter]["Approver"]);
                        // viewIndent.Approval_Status = Convert.ToString(dataSet.Tables[0].Rows[counter]["ApprovalStatus"]);
                        viewIndent.Date = Convert.ToDateTime(dataSet.Tables[0].Rows[counter]["Date"]);
                        viewIndent.LocationId = Convert.ToInt64(dataSet.Tables[0].Rows[counter]["LocationId"]);
                        viewIndent.Location = Convert.ToString(dataSet.Tables[0].Rows[counter]["Location"]);
                        viewIndent.IndentRemarks = Convert.ToString(dataSet.Tables[0].Rows[counter]["Remarks"]);

                        viewIndent.CategoryName = Convert.ToString(dataSet.Tables[0].Rows[counter]["ItemCategoryName"]);
                        viewIndent.ItemCode = Convert.ToString(dataSet.Tables[0].Rows[counter]["ItemCode"]);
                        viewIndent.Units = Convert.ToString(dataSet.Tables[0].Rows[counter]["Unit"]);
                        viewIndent.Description = Convert.ToString(dataSet.Tables[0].Rows[counter]["Description"]);
                        viewIndent.Technical_Specifications = Convert.ToString(dataSet.Tables[0].Rows[counter]["TechnicalSpecification"]);
                        viewIndent.Quantity = Convert.ToInt32(dataSet.Tables[0].Rows[counter]["Quantity"]);
                        // viewIndent.Remarks = Convert.ToString(dataSet.Tables[0].Rows[counter]["Item Remarks"]);

                        viewIndent.Email = Convert.ToString(dataSet.Tables[0].Rows[counter]["Email"]);

                        viewIndents.Add(viewIndent);
                        counter++;
                    }
                    dataSet.Dispose();
                    grid_indentdata.ItemsSource = null;
                    grid_indentdata.ItemsSource = viewIndents;
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void FillApprovalStatusLevel(string poID)
        {
            try
            {
                List<ApprovalStatusLevel> approvalStatusLevels = new List<ApprovalStatusLevel>();
                using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlConnection"].ToString()))
                {
                    connection.Open();
                    SqlCommand testCMD = new SqlCommand("getPOApprovalLevel", connection);
                    testCMD.CommandType = CommandType.StoredProcedure;

                    //testCMD.Parameters.Add(new SqlParameter("@UserId", System.Data.SqlDbType.BigInt, 50) { Value = _login.EmployeeID });
                    testCMD.Parameters.Add(new SqlParameter("@POId", System.Data.SqlDbType.VarChar, 300) { Value = poID });

                    // SqlDataReader dataReader = testCMD.ExecuteReader();
                    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(testCMD);

                    DataSet dataSet = new DataSet();
                    sqlDataAdapter.Fill(dataSet);

                    int counter = 0;

                    while (counter < dataSet.Tables[0].Rows.Count)
                    {
                        ApprovalStatusLevel approvalStatusLevel = new ApprovalStatusLevel();
                        approvalStatusLevel.Approver_Name = Convert.ToString(dataSet.Tables[0].Rows[counter]["ApproverName"]);
                        approvalStatusLevel.Approver_Status = Convert.ToString(dataSet.Tables[0].Rows[counter]["ApprovalStatus"]);
                        approvalStatusLevels.Add(approvalStatusLevel);
                        counter++;
                    }
                    dataSet.Dispose();
                    grid_ApprovalStatusLevel.ItemsSource = null;
                    grid_ApprovalStatusLevel.ItemsSource = approvalStatusLevels;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error occured during Approval Status Check : " + ex.Message,
                   "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private void btn_download_Click_1(object sender, RoutedEventArgs e)
        {
            this.Cursor = Cursors.Wait;
            try
            {
                string targetPath = Convert.ToString(ConfigurationManager.AppSettings["UploadQuotationPath"]);
                string fileName = string.Empty;
                string fileContent = string.Empty;

                try
                {
                    using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlConnection"].ToString()))
                    {
                        connection.Open();
                        SqlCommand testCMD = new SqlCommand("GetQuotationInformation", connection);
                        testCMD.CommandType = CommandType.StoredProcedure;

                        //testCMD.Parameters.Add(new SqlParameter("@UserId", System.Data.SqlDbType.BigInt, 50) { Value = _login.EmployeeID });
                        testCMD.Parameters.Add(new SqlParameter("@IndentID", System.Data.SqlDbType.Int) { Value = Convert.ToInt32(txt_indent_no.Text) });
                        testCMD.Parameters.Add(new SqlParameter("@QuoteNo", System.Data.SqlDbType.Int) { Value = 1 });

                        SqlDataReader dataReader = testCMD.ExecuteReader();
                        while (dataReader.Read())
                        {
                            fileName = dataReader.GetString(0);
                            fileContent = dataReader.GetString(1);
                        }
                        if (fileContent != "")
                        {
                            Byte[] bytes = Convert.FromBase64String(fileContent);

                            if (!Directory.Exists(targetPath))
                            {
                                Directory.CreateDirectory(targetPath);
                            }
                            File.WriteAllBytes(targetPath + "\\" + fileName, bytes);
                            MessageBox.Show("The file has been downloaded to the following path" + targetPath + "\\" + fileName,
                           "Order Management System", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                        else
                            MessageBox.Show("There are no files to download. Please upload a file.",
                          "Order Management System", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error occurred while downloading Quotation information.",
                        "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error occurred while downloading Quotation information.",
                          "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            this.Cursor = null;
        }

        private void btn_download_Click_2(object sender, RoutedEventArgs e)
        {
            this.Cursor = Cursors.Wait;
            try
            {
                string targetPath = Convert.ToString(ConfigurationManager.AppSettings["UploadQuotationPath"]);
                string fileName = string.Empty;
                string fileContent = string.Empty;

                try
                {
                    using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlConnection"].ToString()))
                    {
                        connection.Open();
                        SqlCommand testCMD = new SqlCommand("GetQuotationInformation", connection);
                        testCMD.CommandType = CommandType.StoredProcedure;

                        //testCMD.Parameters.Add(new SqlParameter("@UserId", System.Data.SqlDbType.BigInt, 50) { Value = _login.EmployeeID });
                        testCMD.Parameters.Add(new SqlParameter("@IndentID", System.Data.SqlDbType.Int) { Value = Convert.ToInt32(txt_indent_no.Text) });
                        testCMD.Parameters.Add(new SqlParameter("@QuoteNo", System.Data.SqlDbType.Int) { Value = 2 });

                        SqlDataReader dataReader = testCMD.ExecuteReader();
                        while (dataReader.Read())
                        {
                            fileName = dataReader.GetString(0);
                            fileContent = dataReader.GetString(1);
                        }

                        if (fileContent != "")
                        {
                            Byte[] bytes = Convert.FromBase64String(fileContent);
                            if (!Directory.Exists(targetPath))
                            {
                                Directory.CreateDirectory(targetPath);
                            }
                            File.WriteAllBytes(targetPath + "\\" + fileName, bytes);
                            MessageBox.Show("The file has been downloaded to the following path" + targetPath + "\\" + fileName,
                           "Order Management System", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                        else
                            MessageBox.Show("There are no files to download. Please upload a file.",
                          "Order Management System", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error occurred while downloading Quotation information.",
                        "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error occurred while downloading Quotation information.",
                      "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            this.Cursor = null;
        }

        private void btn_download_Click_3(object sender, RoutedEventArgs e)
        {
            this.Cursor = Cursors.Wait;
            try
            {
                string targetPath = Convert.ToString(ConfigurationManager.AppSettings["UploadQuotationPath"]);
                string fileName = string.Empty;
                string fileContent = string.Empty;

                try
                {
                    using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlConnection"].ToString()))
                    {
                        connection.Open();
                        SqlCommand testCMD = new SqlCommand("GetQuotationInformation", connection);
                        testCMD.CommandType = CommandType.StoredProcedure;

                        //testCMD.Parameters.Add(new SqlParameter("@UserId", System.Data.SqlDbType.BigInt, 50) { Value = _login.EmployeeID });
                        testCMD.Parameters.Add(new SqlParameter("@IndentID", System.Data.SqlDbType.Int) { Value = Convert.ToInt32(txt_indent_no.Text) });
                        testCMD.Parameters.Add(new SqlParameter("@QuoteNo", System.Data.SqlDbType.Int) { Value = 3 });

                        SqlDataReader dataReader = testCMD.ExecuteReader();
                        while (dataReader.Read())
                        {
                            fileName = dataReader.GetString(0);
                            fileContent = dataReader.GetString(1);
                        }

                        if (fileContent != "")
                        {
                            Byte[] bytes = Convert.FromBase64String(fileContent);
                            if (!Directory.Exists(targetPath))
                            {
                                Directory.CreateDirectory(targetPath);
                            }
                            File.WriteAllBytes(targetPath + "\\" + fileName, bytes);
                            MessageBox.Show("The file has been downloaded to the following path" + targetPath + "\\" + fileName,
                           "Order Management System", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                        else
                            MessageBox.Show("There are no files to download. Please upload a file.",
                          "Order Management System", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error occurred while downloading Quotation information.",
                       "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error occurred while downloading Quotation information.",
                       "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            this.Cursor = null;
        }
        private void btn_quotation_upload_Click_1(object sender, RoutedEventArgs e)
        {
            try
            {
                string filePath = string.Empty;
                OpenFileDialog openFileDialog = new OpenFileDialog();
                if (openFileDialog.ShowDialog() == true)
                    filePath = openFileDialog.FileName;


                string fileName = openFileDialog.SafeFileName.Split('.')[0] + "_" + DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Year + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second + "_" + _login.EmployeeID + "." + openFileDialog.SafeFileName.Split('.')[1];
                byte[] fileContent = null;

                //System.IO.FileStream fs = new System.IO.FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                //System.IO.BinaryReader binaryReader = new System.IO.BinaryReader(fs);

                //long byteLength = new System.IO.FileInfo(filePath).Length;
                fileContent = File.ReadAllBytes(filePath);
                String file = Convert.ToBase64String(fileContent);

                //using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                //{
                //    poitems_1 = PullIndentData(stream);
                //}
                //grid_quote_1.ItemsSource = null;
                //grid_quote_1.ItemsSource = poitems_1;
                //File.SetAttributes(filePath, FileAttributes.Normal);
                try
                {
                    using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlConnection"].ToString()))
                    {
                        connection.Open();
                        SqlCommand testCMD = new SqlCommand("InsertQuotationInformation", connection);
                        testCMD.CommandType = CommandType.StoredProcedure;

                        //testCMD.Parameters.Add(new SqlParameter("@UserId", System.Data.SqlDbType.BigInt, 50) { Value = _login.EmployeeID });
                        testCMD.Parameters.Add(new SqlParameter("@QuoteNo", System.Data.SqlDbType.Int, 300) { Value = 1 });
                        testCMD.Parameters.Add(new SqlParameter("@QuotationFileName", System.Data.SqlDbType.VarChar, 300) { Value = fileName });
                        testCMD.Parameters.Add(new SqlParameter("@IndentID", System.Data.SqlDbType.VarChar, 300) { Value = txt_indent_no.Text.ToString() });
                        testCMD.Parameters.Add(new SqlParameter("@PONo", System.Data.SqlDbType.VarChar, 300) { Value = txt_PO_no.Text.ToString() });
                        testCMD.Parameters.Add(new SqlParameter("@FileContent", System.Data.SqlDbType.VarChar) { Value = file });
                        testCMD.Parameters.Add(new SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 300) { Value = _login.EmployeeID });
                        testCMD.Parameters.Add(new SqlParameter("@CreatedDate", System.Data.SqlDbType.DateTime, 300) { Value = DateTime.Now });

                        // SqlDataReader dataReader = testCMD.ExecuteReader();
                        testCMD.ExecuteNonQuery();

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error occurred while saving Quotation information.",
                        "Order Management System", MessageBoxButton.OK, MessageBoxImage.Information);
                }


                //targetPath = targetPath + "\\" + fileName;
                //File.Copy(filePath, targetPath, true);
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

        private void btn_quotation_upload_Click_2(object sender, RoutedEventArgs e)
        {
            try
            {
                string filePath = string.Empty;
                OpenFileDialog openFileDialog = new OpenFileDialog();
                if (openFileDialog.ShowDialog() == true)
                    filePath = openFileDialog.FileName;


                string fileName = openFileDialog.SafeFileName.Split('.')[0] + "_" + DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Year + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second + "_" + _login.EmployeeID + "." + openFileDialog.SafeFileName.Split('.')[1];
                byte[] fileContent = null;

                //System.IO.FileStream fs = new System.IO.FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                //System.IO.BinaryReader binaryReader = new System.IO.BinaryReader(fs);

                //long byteLength = new System.IO.FileInfo(filePath).Length;
                fileContent = File.ReadAllBytes(filePath);
                String file = Convert.ToBase64String(fileContent);

                //using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                //{
                //    poitems_1 = PullIndentData(stream);
                //}
                //grid_quote_1.ItemsSource = null;
                //grid_quote_1.ItemsSource = poitems_1;
                //File.SetAttributes(filePath, FileAttributes.Normal);
                try
                {
                    using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlConnection"].ToString()))
                    {
                        connection.Open();
                        SqlCommand testCMD = new SqlCommand("InsertQuotationInformation", connection);
                        testCMD.CommandType = CommandType.StoredProcedure;

                        //testCMD.Parameters.Add(new SqlParameter("@UserId", System.Data.SqlDbType.BigInt, 50) { Value = _login.EmployeeID });
                        testCMD.Parameters.Add(new SqlParameter("@QuoteNo", System.Data.SqlDbType.Int, 300) { Value = 2 });
                        testCMD.Parameters.Add(new SqlParameter("@QuotationFileName", System.Data.SqlDbType.VarChar, 300) { Value = fileName });
                        testCMD.Parameters.Add(new SqlParameter("@IndentID", System.Data.SqlDbType.VarChar, 300) { Value = txt_indent_no.Text.ToString() });
                        testCMD.Parameters.Add(new SqlParameter("@FileContent", System.Data.SqlDbType.VarChar) { Value = file });
                        testCMD.Parameters.Add(new SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 300) { Value = _login.EmployeeID });
                        testCMD.Parameters.Add(new SqlParameter("@CreatedDate", System.Data.SqlDbType.DateTime, 300) { Value = DateTime.Now });

                        // SqlDataReader dataReader = testCMD.ExecuteReader();
                        testCMD.ExecuteNonQuery();

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error occurred while saving Quotation information.",
                        "Order Management System", MessageBoxButton.OK, MessageBoxImage.Information);
                }


                //targetPath = targetPath + "\\" + fileName;
                //File.Copy(filePath, targetPath, true);
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
        private void btn_quotation_upload_Click_3(object sender, RoutedEventArgs e)
        {
            try
            {
                string filePath = string.Empty;
                OpenFileDialog openFileDialog = new OpenFileDialog();
                if (openFileDialog.ShowDialog() == true)
                    filePath = openFileDialog.FileName;


                string fileName = openFileDialog.SafeFileName.Split('.')[0] + "_" + DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Year + DateTime.Now.Hour + DateTime.Now.Minute + DateTime.Now.Second + "_" + _login.EmployeeID + "." + openFileDialog.SafeFileName.Split('.')[1];
                byte[] fileContent = null;

                //System.IO.FileStream fs = new System.IO.FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);
                //System.IO.BinaryReader binaryReader = new System.IO.BinaryReader(fs);

                //long byteLength = new System.IO.FileInfo(filePath).Length;
                fileContent = File.ReadAllBytes(filePath);
                String file = Convert.ToBase64String(fileContent);

                //using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                //{
                //    poitems_1 = PullIndentData(stream);
                //}
                //grid_quote_1.ItemsSource = null;
                //grid_quote_1.ItemsSource = poitems_1;
                //File.SetAttributes(filePath, FileAttributes.Normal);
                try
                {
                    using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlConnection"].ToString()))
                    {
                        connection.Open();
                        SqlCommand testCMD = new SqlCommand("InsertQuotationInformation", connection);
                        testCMD.CommandType = CommandType.StoredProcedure;

                        //testCMD.Parameters.Add(new SqlParameter("@UserId", System.Data.SqlDbType.BigInt, 50) { Value = _login.EmployeeID });
                        testCMD.Parameters.Add(new SqlParameter("@QuoteNo", System.Data.SqlDbType.Int, 300) { Value = 3 });
                        testCMD.Parameters.Add(new SqlParameter("@QuotationFileName", System.Data.SqlDbType.VarChar, 300) { Value = fileName });
                        testCMD.Parameters.Add(new SqlParameter("@IndentID", System.Data.SqlDbType.VarChar, 300) { Value = txt_indent_no.Text.ToString() });
                        testCMD.Parameters.Add(new SqlParameter("@FileContent", System.Data.SqlDbType.VarChar) { Value = file });
                        testCMD.Parameters.Add(new SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 300) { Value = _login.EmployeeID });
                        testCMD.Parameters.Add(new SqlParameter("@CreatedDate", System.Data.SqlDbType.DateTime, 300) { Value = DateTime.Now });

                        // SqlDataReader dataReader = testCMD.ExecuteReader();
                        testCMD.ExecuteNonQuery();

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error occurred while saving Quotation information.",
                        "Order Management System", MessageBoxButton.OK, MessageBoxImage.Information);
                }


                //targetPath = targetPath + "\\" + fileName;
                //File.Copy(filePath, targetPath, true);
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
                                        if (!Convert.ToString(reader.GetValue(column)).ToLower().Contains("s.no"))
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
                                    poitem.Offer_Number = lst[1];
                                    poitem.Vendor_Name = lst[2];
                                    poitem.Vendor_Code = lst[3];
                                    poitem.Offer_Date = lst[4];
                                    poitem.Contact_No = lst[5];
                                    poitem.Contact_Person = lst[6];
                                    poitem.Description = lst[7];
                                    poitem.Quantity = Convert.ToInt32(lst[8]);
                                    poitem.Units = lst[9];
                                    poitem.Unit_Price = Convert.ToDouble(lst[10]);
                                    poitem.HSN = Convert.ToInt32(lst[11]);
                                    poitem.GST_Value = Convert.ToDecimal(lst[12]);
                                    poitem.Total_Price = Convert.ToDecimal(lst[13]);

                                    poitems.Add(poitem);
                                    // loadExcelIndent.ExcelIndents = lstGridIndent;
                                }
                            }
                            _readData = new List<List<string>>();
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

        private void DisableFields()
        {
            txt_indent_no.IsEnabled = false;
            checkbox_Approve1.IsEnabled = false;
            checkbox_Approve2.IsEnabled = false;
            checkbox_Approve3.IsEnabled = false;
            btn_upload_1.IsEnabled = false;
            btn_upload_2.IsEnabled = false;
            btn_upload_3.IsEnabled = false;
            btn_quotation_upload_1.IsEnabled = false;
            btn_quotation_upload_2.IsEnabled = false;
            btn_quotation_upload_3.IsEnabled = false;
            txt_PO_Remarks.IsEnabled = false;
            cbx_ApprovalStatus_id.IsEnabled = false;
            btn_Approve.IsEnabled = false;
            btn_Create_PO.IsEnabled = false;
            // btn_Generate.IsEnabled = false;
        }

        private void EnableFields()
        {
            txt_indent_no.IsEnabled = true;
            checkbox_Approve1.IsEnabled = true;
            checkbox_Approve2.IsEnabled = true;
            checkbox_Approve3.IsEnabled = true;
            btn_upload_1.IsEnabled = true;
            btn_upload_2.IsEnabled = true;
            btn_upload_3.IsEnabled = true;
            btn_quotation_upload_1.IsEnabled = true;
            btn_quotation_upload_2.IsEnabled = true;
            btn_quotation_upload_3.IsEnabled = true;
            txt_PO_Remarks.IsEnabled = true;
            cbx_ApprovalStatus_id.IsEnabled = true;
            checkbox_Approve1.IsEnabled = true;
            checkbox_Approve2.IsEnabled = true;
            checkbox_Approve2.IsEnabled = true;

            btn_Create_PO.IsEnabled = true;
            // btn_Generate.IsEnabled = true;
        }


        private void grid_quote_3_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void grid_quote_2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void checkbox_Approve1_Checked(object sender, RoutedEventArgs e)
        {
            if (poitems_1.Count > 0)
            {
                checkbox_Approve2.IsChecked = false;
                checkbox_Approve3.IsChecked = false;
                grid_po_confirmation.ItemsSource = null;
                grid_po_confirmation.ItemsSource = poitems_1;
                poitems_4 = null;
                poitems_4 = poitems_1;
                btn_Approve.IsEnabled = true;
                txt_GST.Text = poitems_4[0].GST_Value.ToString();
            }
            else
            {
                checkbox_Approve1.IsChecked = false;
            }

        }
        private void checkbox_Approve1_Unchecked(object sender, RoutedEventArgs e)
        {
            grid_po_confirmation.ItemsSource = null;
        }

        private void checkbox_Approve2_Checked(object sender, RoutedEventArgs e)
        {
            if (poitems_2.Count > 0)
            {
                checkbox_Approve1.IsChecked = false;
                checkbox_Approve3.IsChecked = false;
                grid_po_confirmation.ItemsSource = null;
                grid_po_confirmation.ItemsSource = poitems_2;
                poitems_4 = null;
                poitems_4 = poitems_1;
                btn_Approve.IsEnabled = true;
                txt_GST.Text = poitems_4[0].GST_Value.ToString();
            }
            else
            {
                checkbox_Approve2.IsChecked = false;
            }
        }

        private void checkbox_Approve2_Unchecked(object sender, RoutedEventArgs e)
        {
            grid_po_confirmation.ItemsSource = null;
        }

        private void checkbox_Approve3_Checked(object sender, RoutedEventArgs e)
        {
            if (poitems_3.Count > 0)
            {
                checkbox_Approve1.IsChecked = false;
                checkbox_Approve2.IsChecked = false;
                grid_po_confirmation.ItemsSource = null;
                grid_po_confirmation.ItemsSource = poitems_3;
                poitems_4 = null;
                poitems_4 = poitems_1;
                btn_Approve.IsEnabled = true;
                txt_GST.Text = poitems_4[0].GST_Value.ToString();
            }
            else
            {
                checkbox_Approve3.IsChecked = false;
            }
        }

        private void checkbox_Approve3_Unchecked(object sender, RoutedEventArgs e)
        {
            grid_po_confirmation.ItemsSource = null;
        }

        private void grid_indentdata_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        protected override void OnClosing(CancelEventArgs e)
        {
            Menu menu = new Menu(_login);
            menu.Show();
        }

        private void txt_PO_Remarks_TextChanged(object sender, TextChangedEventArgs e)
        {

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
        private void GeneratePO()
        {
            try
            {
                this.Cursor = Cursors.Wait;
                indentNo = Convert.ToInt64(txt_indent_no.Text);
                Dictionary<string, string> _headers = GetHeaders();
                ExportPO poData = new ExportPO();
                List<Poitem> poItems = new List<Poitem>();
                LocationAddress _locationAddress = (from loc in orderManagementContext.LocationAddress
                                                    join ind in orderManagementContext.IndentMaster
                                                    on loc.LocationCodeId equals ind.LocationCodeId
                                                    where ind.IndentId == indentNo
                                                    select loc).FirstOrDefault();
                var locationDetails = (from loc in orderManagementContext.LocationCode
                                       join ind in orderManagementContext.IndentMaster
                                       on loc.LocationCodeId equals ind.LocationCodeId
                                       where ind.IndentId == indentNo
                                       select loc).FirstOrDefault();

                string locationName = "";
                if (locationDetails != null)
                {
                    locationName = locationDetails.LocationId.ToString().Trim() + " - " + locationDetails.LocationName.Trim();
                }
                poData.IndentID = indentNo;
                poData.PODate = DateTime.Now;
                poData.POID = Convert.ToInt32(txt_PO_no.Text);
                poData.Remarks = txt_PO_Remarks.Text;
                poData.Email = _login.UserEmail;
                poData.Poitems = poitems_4;
                poData.LocatioName = locationName;
                poData.LocationAddressInfo = _locationAddress;
                GeneratePurchaseOrder(_headers, poData);
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error Occured during PO Generation" + ex.Message,
                   "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            this.Cursor = null;
        }
        private void btn_generate_Click(object sender, RoutedEventArgs e)
        {
            if (poitems_1.Count == 0 && poitems_2.Count == 0 && poitems_3.Count == 0)
            {
                MessageBox.Show("Please upload any one Quotation",
                      "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                this.Cursor = null;
                return;
            }
            if (poitems_4.Count == 0)
            {
                MessageBox.Show("Please approve any of the 3 Quotations",
                     "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                this.Cursor = null;
                return;
            }
            if (txt_indent_no.Text == null || txt_indent_no.Text == "")
            {
                MessageBox.Show("Please enter indent number",
                     "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                this.Cursor = null;
                return;
            }
            if (txt_PO_no.Text == null || txt_PO_no.Text == "")
            {
                MessageBox.Show("Please Select a generated PO",
                     "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                this.Cursor = null;
                return;
            }
            GeneratePO();
        }
        private void btn_Create_PO_Click(object sender, RoutedEventArgs e)
        {
            this.Cursor = Cursors.Wait;
            try
            {
                bool isMailSent = false;
                if (poitems_1.Count == 0 && poitems_2.Count == 0 && poitems_3.Count == 0)
                {
                    MessageBox.Show("Please upload any one Quotation",
                          "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                    this.Cursor = null;
                    return;
                }
                //if(poitems_4.Count==0)
                //{
                //    MessageBox.Show("Please approve any of the 3 Quotations",
                //         "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                //    this.Cursor = null;
                //    return;
                //}
                if (txt_indent_no.Text == null || txt_indent_no.Text == "")
                {
                    MessageBox.Show("Please enter indent number",
                         "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                    this.Cursor = null;
                    return;
                }
                using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlConnection"].ToString()))
                {
                    connection.Open();
                    SqlCommand testCMD = new SqlCommand("CreatePurchaseOrder", connection);
                    testCMD.CommandType = CommandType.StoredProcedure;

                    //testCMD.Parameters.Add(new SqlParameter("@Date", System.Data.SqlDbType.VarChar, 50) { Value = saveIndent.Date });
                    testCMD.Parameters.Add(new SqlParameter("@IndentId", System.Data.SqlDbType.BigInt, 50) { Value = txt_indent_no.Text });
                    testCMD.Parameters.Add(new SqlParameter("@Remarks", System.Data.SqlDbType.VarChar, 300) { Value = txt_PO_Remarks.Text });
                    testCMD.Parameters.Add(new SqlParameter("@CreatedBy", System.Data.SqlDbType.BigInt, 50) { Value = _login.EmployeeID });
                    testCMD.Parameters.Add(new SqlParameter("@POId", System.Data.SqlDbType.BigInt, 50) { Value = poID });
                    testCMD.Parameters["@POId"].Direction = ParameterDirection.Output;

                    testCMD.ExecuteNonQuery(); // read output value from @NewId 
                    poID = Convert.ToInt32(testCMD.Parameters["@POId"].Value);

                    if (poitems_1.Count > 0)
                    {
                        foreach (var i in poitems_1)
                        {
                            SqlCommand testCMD1 = new SqlCommand("CreatePurchaseOrderDetails", connection);
                            testCMD1.CommandType = CommandType.StoredProcedure;
                            testCMD1.Parameters.Add(new SqlParameter("@POId", System.Data.SqlDbType.BigInt, 50) { Value = poID });
                            testCMD1.Parameters.Add(new SqlParameter("@Offer_Number", System.Data.SqlDbType.VarChar, 300) { Value = i.Offer_Number.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Vendor_Name", System.Data.SqlDbType.VarChar, 300) { Value = i.Vendor_Name.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Vendor_Code", System.Data.SqlDbType.VarChar, 300) { Value = i.Vendor_Code.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Offer_Date", System.Data.SqlDbType.DateTime, 300) { Value = i.Offer_Date.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Contact_No", System.Data.SqlDbType.NChar, 20) { Value = i.Contact_No.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Contact_Person", System.Data.SqlDbType.VarChar, 300) { Value = i.Contact_Person.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Description", System.Data.SqlDbType.VarChar, 300) { Value = i.Description.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Quantity", System.Data.SqlDbType.BigInt, 50) { Value = i.Quantity });
                            testCMD1.Parameters.Add(new SqlParameter("@Units", System.Data.SqlDbType.VarChar, 50) { Value = i.Units.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@UnitPrice", System.Data.SqlDbType.Decimal, 50) { Value = i.Unit_Price });
                            testCMD1.Parameters.Add(new SqlParameter("@TotalPrice", System.Data.SqlDbType.Decimal, 50) { Value = i.Total_Price });
                            testCMD1.Parameters.Add(new SqlParameter("@HSN", System.Data.SqlDbType.Decimal, 50) { Value = i.HSN });
                            testCMD1.Parameters.Add(new SqlParameter("@GST_Value", System.Data.SqlDbType.Decimal, 50) { Value = i.GST_Value });
                            testCMD1.Parameters.Add(new SqlParameter("@QuoteNo", System.Data.SqlDbType.Int, 50) { Value = 1 });
                            testCMD1.Parameters.Add(new SqlParameter("@CreatedBy", System.Data.SqlDbType.BigInt, 50) { Value = _login.EmployeeID });
                            testCMD1.ExecuteNonQuery(); // read output value from @NewId 
                        }
                    }

                    if (poitems_2.Count > 0)
                    {
                        foreach (var i in poitems_2)
                        {
                            SqlCommand testCMD1 = new SqlCommand("CreatePurchaseOrderDetails", connection);
                            testCMD1.CommandType = CommandType.StoredProcedure;
                            testCMD1.Parameters.Add(new SqlParameter("@POId", System.Data.SqlDbType.BigInt, 50) { Value = poID });
                            testCMD1.Parameters.Add(new SqlParameter("@Offer_Number", System.Data.SqlDbType.VarChar, 300) { Value = i.Offer_Number.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Vendor_Name", System.Data.SqlDbType.VarChar, 300) { Value = i.Vendor_Name.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Vendor_Code", System.Data.SqlDbType.VarChar, 300) { Value = i.Vendor_Code.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Offer_Date", System.Data.SqlDbType.DateTime, 300) { Value = i.Offer_Date.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Contact_No", System.Data.SqlDbType.NChar, 20) { Value = i.Contact_No.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Contact_Person", System.Data.SqlDbType.VarChar, 300) { Value = i.Contact_Person.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Description", System.Data.SqlDbType.VarChar, 300) { Value = i.Description.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Quantity", System.Data.SqlDbType.BigInt, 50) { Value = i.Quantity });
                            testCMD1.Parameters.Add(new SqlParameter("@Units", System.Data.SqlDbType.VarChar, 50) { Value = i.Units.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@UnitPrice", System.Data.SqlDbType.Decimal, 50) { Value = i.Unit_Price });
                            testCMD1.Parameters.Add(new SqlParameter("@TotalPrice", System.Data.SqlDbType.Decimal, 50) { Value = i.Total_Price });
                            testCMD1.Parameters.Add(new SqlParameter("@QuoteNo", System.Data.SqlDbType.Int, 50) { Value = 2 });
                            testCMD1.Parameters.Add(new SqlParameter("@HSN", System.Data.SqlDbType.Decimal, 50) { Value = i.HSN });
                            testCMD1.Parameters.Add(new SqlParameter("@GST_Value", System.Data.SqlDbType.Decimal, 50) { Value = i.GST_Value });
                            testCMD1.Parameters.Add(new SqlParameter("@CreatedBy", System.Data.SqlDbType.BigInt, 50) { Value = _login.EmployeeID });

                            testCMD1.ExecuteNonQuery(); // read output value from @NewId 
                        }
                    }

                    if (poitems_3.Count > 0)
                    {
                        foreach (var i in poitems_3)
                        {
                            SqlCommand testCMD1 = new SqlCommand("CreatePurchaseOrderDetails", connection);
                            testCMD1.CommandType = CommandType.StoredProcedure;
                            testCMD1.Parameters.Add(new SqlParameter("@POId", System.Data.SqlDbType.BigInt, 50) { Value = poID });
                            testCMD1.Parameters.Add(new SqlParameter("@Offer_Number", System.Data.SqlDbType.VarChar, 300) { Value = i.Offer_Number.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Vendor_Name", System.Data.SqlDbType.VarChar, 300) { Value = i.Vendor_Name.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Vendor_Code", System.Data.SqlDbType.VarChar, 300) { Value = i.Vendor_Code.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Offer_Date", System.Data.SqlDbType.DateTime, 300) { Value = i.Offer_Date.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Contact_No", System.Data.SqlDbType.NChar, 20) { Value = i.Contact_No.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Contact_Person", System.Data.SqlDbType.VarChar, 300) { Value = i.Contact_Person.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Description", System.Data.SqlDbType.VarChar, 300) { Value = i.Description.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Quantity", System.Data.SqlDbType.BigInt, 50) { Value = i.Quantity });
                            testCMD1.Parameters.Add(new SqlParameter("@Units", System.Data.SqlDbType.VarChar, 50) { Value = i.Units.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@UnitPrice", System.Data.SqlDbType.Decimal, 50) { Value = i.Unit_Price });
                            testCMD1.Parameters.Add(new SqlParameter("@TotalPrice", System.Data.SqlDbType.Decimal, 50) { Value = i.Total_Price });
                            testCMD1.Parameters.Add(new SqlParameter("@QuoteNo", System.Data.SqlDbType.Int, 50) { Value = 3 });
                            testCMD1.Parameters.Add(new SqlParameter("@HSN", System.Data.SqlDbType.Decimal, 50) { Value = i.HSN });
                            testCMD1.Parameters.Add(new SqlParameter("@GST_Value", System.Data.SqlDbType.Decimal, 50) { Value = i.GST_Value });
                            testCMD1.Parameters.Add(new SqlParameter("@CreatedBy", System.Data.SqlDbType.BigInt, 50) { Value = _login.EmployeeID });
                            testCMD1.ExecuteNonQuery(); // read output value from @NewId 
                        }
                    }

                    SqlCommand testCMD2 = new SqlCommand("CreatePurchaseOrderApproval", connection);
                    testCMD2.CommandType = CommandType.StoredProcedure;
                    testCMD2.Parameters.Add(new SqlParameter("@POID", System.Data.SqlDbType.BigInt, 50) { Value = poID });
                    //testCMD2.Parameters.Add(new SqlParameter("@ApprovalID", System.Data.SqlDbType.BigInt, 50) { Value = saveIndent.ApprovalID });
                    testCMD2.Parameters.Add(new SqlParameter("@ApprovalStatusID", System.Data.SqlDbType.BigInt, 50) { Value = 1 });
                    // testCMD2.Parameters.Add(new SqlParameter("@CreatedDate", System.Data.SqlDbType.DateTime, 50) { Value = saveIndent.CreateDate });
                    testCMD2.Parameters.Add(new SqlParameter("@CreatedBy", System.Data.SqlDbType.BigInt, 50) { Value = _login.EmployeeID });
                    testCMD2.ExecuteNonQuery();


                    mailTo = (from a in orderManagementContext.UserMaster
                              where a.EmployeeId == (from emp in orderManagementContext.Employee where emp.EmployeeId == _login.EmployeeID select emp.ReportsTo).FirstOrDefault()
                              select a.Email).FirstOrDefault();

                    string body = GeneratePOMail(poID.ToString());
                    isMailSent = SendMail(body, "Purchase Order " + poID + " is created by " + _login.UserName + ". Kindly Approve or Deny the Purchase Order ");
                    if (isMailSent)
                    {
                        MessageBox.Show("Purchase Order " + poID + " is created.", "Order Management System",
              MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        MessageBox.Show("Purchase Order " + poID + " is created. But mail is not sent", "Order Management System",
             MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    txt_PO_no.Text = poID.ToString();
                    //GeneratePO();
                    DisableFields();
                    DisableUploadButtons();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Occured during PO Creation." + ex.Message,
                      "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            this.Cursor = null;
        }

        private string GeneratePOMail(string PO_Id)
        {
            try
            {
                int counter = 1;
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
                htmlString.Append("<table width='100%' border=1 border-style='solid' border-color='#8CBD48'>");
                htmlString.Append("<tr>");
                htmlString.Append("<td colspan=7 style='background-color:#ffffff;color:#6C6C6C;font-family:Verdana;'><b>Purchase Order : " + poID + "</b></td>");
                htmlString.Append("</table>");
                htmlString.Append("<table width='100%' border=1 border-style='solid' border-color='#8CBD48'>");
                htmlString.Append("<tr>");
                htmlString.Append("<td colspan=7 style='background-color:#EBF8F0;color:#6C6C6C;font-family:Verdana;'>Quotation : 1</td>");
                htmlString.Append("</table>");
                //htmlString.Append("<br/>");
                if (poitems_1.Count > 0)
                {
                    htmlString.Append("<table width='100%'  border=1 border-style='solid' border-color='#8CBD48'>");
                    htmlString.Append("<tr style='background-color:#609F19;color:#FFFFFF;font-family:Verdana;font-size:11;vertical-align:middle'><td>S.No</td>");
                    htmlString.Append("<td>Offer No</td>");
                    htmlString.Append("<td>Vendor Name</td>");
                    htmlString.Append("<td>Vendor Code</td>");
                    htmlString.Append("<td>Offer Date</td>");
                    htmlString.Append("<td>Contact No</td>");
                    htmlString.Append("<td>Contact Person</td>");
                    htmlString.Append("<td>Description</td>");
                    htmlString.Append("<td>Qty</td>");
                    htmlString.Append("<td>Units</td>");
                    htmlString.Append("<td>Unit Price</td>");
                    htmlString.Append("<td>HSN</td>");
                    htmlString.Append("<td>GST %</td>");
                    htmlString.Append("<td>Total Price</td>");
                    htmlString.Append("</tr>");
                    counter = 1;
                    for (int k = 0; k < poitems_1.Count; k++)
                    {
                        if (k <= poitems_1.Count - 1)
                        {
                            htmlString.Append("<tr style='background-color:#FFFFFF;color:#6C6C6C;font-family:Verdana;font-size:11;'>");
                            htmlString.Append("<td>" + counter + "</td>");
                            htmlString.Append("<td>" + poitems_1[k].Offer_Number + "</td>");
                            htmlString.Append("<td>" + poitems_1[k].Vendor_Name + "</td>");
                            htmlString.Append("<td>" + poitems_1[k].Vendor_Code + "</td>");
                            htmlString.Append("<td>" + poitems_1[k].Offer_Date + "</td>");
                            htmlString.Append("<td>" + poitems_1[k].Contact_No + "</td>");
                            htmlString.Append("<td>" + poitems_1[k].Contact_Person + "</td>");
                            htmlString.Append("<td>" + poitems_1[k].Description + "</td>");
                            htmlString.Append("<td>" + poitems_1[k].Quantity + "</td>");
                            htmlString.Append("<td>" + poitems_1[k].Units + "</td>");
                            htmlString.Append("<td>" + poitems_1[k].Unit_Price + "</td>");
                            htmlString.Append("<td>" + poitems_1[k].HSN + "</td>");
                            htmlString.Append("<td>" + poitems_1[k].GST_Value + "</td>");
                            htmlString.Append("<td>" + poitems_1[k].Total_Price + "</td>");

                            htmlString.Append("</tr>");
                            counter++;
                        }
                    }
                    htmlString.Append("</table>");
                }
                htmlString.Append("<br/>");

                if (poitems_2.Count > 0)
                {
                    htmlString.Append("<table width='100%'  border=1 border-style='solid' border-color='#8CBD48'>");
                    htmlString.Append("<tr>");
                    htmlString.Append("<td colspan=7 style='background-color:#EBF8F0;color:#6C6C6C;font-family:Verdana;'>Quotation : 2</td>");
                    //htmlString.Append("<td  style='background-color:#EBF8F0;color:#6C6C6C;font-family:Verdana;'><i>Indent No :</td>");
                    //htmlString.Append("<td  style='background-color:#EBF8F0;color:#6C6C6C;font-family:Verdana;'><i>" + indentNo + "</i></td></tr>");
                    htmlString.Append("</table>");
                    htmlString.Append("<table width='100%'  border=1 border-style='solid' border-color='#8CBD48'>");
                    htmlString.Append("<tr style='background-color:#609F19;color:#FFFFFF;font-family:Verdana;font-size:11;vertical-align:middle'><td>S.No</td>");
                    htmlString.Append("<td>Offer No</td>");
                    htmlString.Append("<td>Vendor Name</td>");
                    htmlString.Append("<td>Vendor Code</td>");
                    htmlString.Append("<td>Offer Date</td>");
                    htmlString.Append("<td>Contact No</td>");
                    htmlString.Append("<td>Contact Person</td>");
                    htmlString.Append("<td>Description</td>");
                    htmlString.Append("<td>Qty</td>");
                    htmlString.Append("<td>Units</td>");
                    htmlString.Append("<td>Unit Price</td>");
                    htmlString.Append("<td>HSN</td>");
                    htmlString.Append("<td>GST %</td>");
                    htmlString.Append("<td>Total Price</td>");
                    htmlString.Append("</tr>");
                    counter = 1;
                    for (int k = 0; k < poitems_2.Count; k++)
                    {
                        if (k <= poitems_2.Count - 1)
                        {
                            htmlString.Append("<tr style='background-color:#FFFFFF;color:#6C6C6C;font-family:Verdana;font-size:11;'>");
                            htmlString.Append("<td>" + counter + "</td>");
                            htmlString.Append("<td>" + poitems_2[k].Offer_Number + "</td>");
                            htmlString.Append("<td>" + poitems_2[k].Vendor_Name + "</td>");
                            htmlString.Append("<td>" + poitems_2[k].Vendor_Code + "</td>");
                            htmlString.Append("<td>" + poitems_2[k].Offer_Date + "</td>");
                            htmlString.Append("<td>" + poitems_2[k].Contact_No + "</td>");
                            htmlString.Append("<td>" + poitems_2[k].Contact_Person + "</td>");
                            htmlString.Append("<td>" + poitems_2[k].Description + "</td>");
                            htmlString.Append("<td>" + poitems_2[k].Quantity + "</td>");
                            htmlString.Append("<td>" + poitems_2[k].Units + "</td>");
                            htmlString.Append("<td>" + poitems_2[k].Unit_Price + "</td>");
                            htmlString.Append("<td>" + poitems_2[k].HSN + "</td>");
                            htmlString.Append("<td>" + poitems_2[k].GST_Value + "</td>");
                            htmlString.Append("<td>" + poitems_2[k].Total_Price + "</td>");

                            htmlString.Append("</tr>");
                            counter++;
                        }
                    }
                    htmlString.Append("<tr>");
                    htmlString.Append("</table>");
                }

                htmlString.Append("<br/>");

                if (poitems_3.Count > 0)
                {
                    htmlString.Append("<table width='100%'  border=1 border-style='solid' border-color='#8CBD48'>");
                    htmlString.Append("<tr>");
                    htmlString.Append("<td colspan=7 style='background-color:#EBF8F0;color:#6C6C6C;font-family:Verdana;'>Quotation : 3</td>");
                    //htmlString.Append("<td  style='background-color:#EBF8F0;color:#6C6C6C;font-family:Verdana;'><i>Indent No :</td>");
                    //htmlString.Append("<td  style='background-color:#EBF8F0;color:#6C6C6C;font-family:Verdana;'><i>" + indentNo + "</i></td></tr>");
                    htmlString.Append("</table>");
                    htmlString.Append("<table width='100%'  border=1 border-style='solid' border-color='#8CBD48'>");
                    htmlString.Append("<tr style='background-color:#609F19;color:#FFFFFF;font-family:Verdana;font-size:11;vertical-align:middle'><td>S.No</td>");
                    htmlString.Append("<td>Offer No</td>");
                    htmlString.Append("<td>Vendor Name</td>");
                    htmlString.Append("<td>Vendor Code</td>");
                    htmlString.Append("<td>Offer Date</td>");
                    htmlString.Append("<td>Contact No</td>");
                    htmlString.Append("<td>Contact Person</td>");
                    htmlString.Append("<td>Description</td>");
                    htmlString.Append("<td>Qty</td>");
                    htmlString.Append("<td>Units</td>");
                    htmlString.Append("<td>Unit Price</td>");
                    htmlString.Append("<td>HSN</td>");
                    htmlString.Append("<td>GST %</td>");
                    htmlString.Append("<td>Total Price</td>");
                    htmlString.Append("</tr>");
                    counter = 1;
                    for (int k = 0; k < poitems_3.Count; k++)
                    {
                        if (k <= poitems_3.Count - 1)
                        {
                            htmlString.Append("<tr style='background-color:#FFFFFF;color:#6C6C6C;font-family:Verdana;font-size:11;'>");
                            htmlString.Append("<td>" + counter + "</td>");
                            htmlString.Append("<td>" + poitems_3[k].Offer_Number + "</td>");
                            htmlString.Append("<td>" + poitems_3[k].Vendor_Name + "</td>");
                            htmlString.Append("<td>" + poitems_3[k].Vendor_Code + "</td>");
                            htmlString.Append("<td>" + poitems_3[k].Offer_Date + "</td>");
                            htmlString.Append("<td>" + poitems_3[k].Contact_No + "</td>");
                            htmlString.Append("<td>" + poitems_3[k].Contact_Person + "</td>");
                            htmlString.Append("<td>" + poitems_3[k].Description + "</td>");
                            htmlString.Append("<td>" + poitems_3[k].Quantity + "</td>");
                            htmlString.Append("<td>" + poitems_3[k].Units + "</td>");
                            htmlString.Append("<td>" + poitems_3[k].Unit_Price + "</td>");
                            htmlString.Append("<td>" + poitems_3[k].HSN + "</td>");
                            htmlString.Append("<td>" + poitems_3[k].GST_Value + "</td>");
                            htmlString.Append("<td>" + poitems_3[k].Total_Price + "</td>");

                            htmlString.Append("</tr>");
                            counter++;
                        }
                    }
                    htmlString.Append("<tr>");

                    htmlString.Append("</table>");
                }
                htmlString.Append("<br/>");

                if (poitems_4.Count > 0)
                {
                    htmlString.Append("<table width='100%'  border=1 border-style='solid' border-color='#8CBD48'>");
                    htmlString.Append("<tr>");
                    htmlString.Append("<td colspan=7 style='background-color:#EBF8F0;color:#ff0000;font-family:Verdana;'>Approved Quotation</td>");
                    //htmlString.Append("<td  style='background-color:#EBF8F0;color:#6C6C6C;font-family:Verdana;'><i>Indent No :</td>");
                    //htmlString.Append("<td  style='background-color:#EBF8F0;color:#6C6C6C;font-family:Verdana;'><i>" + indentNo + "</i></td></tr>");
                    htmlString.Append("</table>");
                    htmlString.Append("<table width='100%'  border=1 border-style='solid' border-color='#8CBD48'>");
                    htmlString.Append("<tr style='background-color:#609F19;color:#FFFFFF;font-family:Verdana;font-size:11;vertical-align:middle'><td>S.No</td>");
                    htmlString.Append("<td>Offer No</td>");
                    htmlString.Append("<td>Vendor Name</td>");
                    htmlString.Append("<td>Vendor Code</td>");
                    htmlString.Append("<td>Offer Date</td>");
                    htmlString.Append("<td>Contact No</td>");
                    htmlString.Append("<td>Contact Person</td>");
                    htmlString.Append("<td>Description</td>");
                    htmlString.Append("<td>Qty</td>");
                    htmlString.Append("<td>Units</td>");
                    htmlString.Append("<td>Unit Price</td>");
                    htmlString.Append("<td>HSN</td>");
                    htmlString.Append("<td>GST %</td>");
                    htmlString.Append("<td>Total Price</td>");
                    htmlString.Append("</tr>");
                    counter = 1;
                    for (int k = 0; k < poitems_4.Count; k++)
                    {
                        if (k <= poitems_4.Count - 1)
                        {
                            htmlString.Append("<tr style='background-color:#FFFFFF;color:#6C6C6C;font-family:Verdana;font-size:11;'>");
                            htmlString.Append("<td>" + counter + "</td>");
                            htmlString.Append("<td>" + poitems_4[k].Offer_Number + "</td>");
                            htmlString.Append("<td>" + poitems_4[k].Vendor_Name + "</td>");
                            htmlString.Append("<td>" + poitems_4[k].Vendor_Code + "</td>");
                            htmlString.Append("<td>" + poitems_4[k].Offer_Date + "</td>");
                            htmlString.Append("<td>" + poitems_4[k].Contact_No + "</td>");
                            htmlString.Append("<td>" + poitems_4[k].Contact_Person + "</td>");
                            htmlString.Append("<td>" + poitems_4[k].Description + "</td>");
                            htmlString.Append("<td>" + poitems_4[k].Quantity + "</td>");
                            htmlString.Append("<td>" + poitems_4[k].Units + "</td>");
                            htmlString.Append("<td>" + poitems_4[k].Unit_Price + "</td>");
                            htmlString.Append("<td>" + poitems_4[k].HSN + "</td>");
                            htmlString.Append("<td>" + poitems_4[k].GST_Value + "</td>");
                            htmlString.Append("<td>" + poitems_4[k].Total_Price + "</td>");

                            htmlString.Append("</tr>");
                            counter++;
                        }
                    }
                    htmlString.Append("<tr>");
                    htmlString.Append("</table>");
                }
                return htmlString.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred. Please contact the administrator", "Order Management System",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }

        }

        private void cbx_ApprovalStatus_id_DropDownOpened(object sender, EventArgs e)
        {
            try
            {
                if (txt_PO_no.Text != "")
                {
                    btn_Approve.IsEnabled = true;
                }
                else
                {
                    LoadApprovalStatus();
                }
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
        private bool SendMail(string body, string content)
        {
            bool mailSent = false;
            try
            {
                string message = DateTime.Now + " In SendMail\n";

                using (MailMessage mm = new MailMessage())
                {
                    mm.From = new MailAddress(Convert.ToString(ConfigurationManager.AppSettings["MailFrom"]));

                    mm.To.Add(mailTo);

                    mm.Subject = "Purchase Order - " + poID.ToString();
                    body = body + content + ConfigurationManager.AppSettings["url"] + "/" + poID.ToString() + "?type=P";
                    mm.Body = body;

                    mm.IsBodyHtml = true;

                    // mm.Attachments.Add(new System.Net.Mail.Attachment(filePathLocation));

                    SmtpClient smtp = new SmtpClient();
                    smtp.Host = ConfigurationManager.AppSettings["Host"];
                    smtp.EnableSsl = false;
                    NetworkCredential NetworkCred = new NetworkCredential(ConfigurationManager.AppSettings["Username"], ConfigurationManager.AppSettings["Password"]);
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



        private void btn_Approve_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                this.Cursor = Cursors.Wait;
                poID = Convert.ToInt64(txt_PO_no.Text);
                indentNo = Convert.ToInt64(txt_indent_no.Text);
                int approvalStatus = Convert.ToInt32(cbx_ApprovalStatus_id.SelectedValue);
                bool isMailSent = false;
                if (poitems_4.Count < 1)
                {
                    MessageBox.Show("Please select any one Quotation ", "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlConnection"].ToString()))
                    {
                        connection.Open();

                        SqlCommand testCMD = new SqlCommand("DeletePurchaseOrderDetails", connection);
                        testCMD.CommandType = CommandType.StoredProcedure;
                        testCMD.Parameters.Add(new SqlParameter("@POId", System.Data.SqlDbType.BigInt, 50) { Value = poID });
                        testCMD.Parameters.Add(new SqlParameter("@QNo", System.Data.SqlDbType.BigInt, 50) { Value = 4 });
                        testCMD.ExecuteNonQuery();

                        foreach (var i in poitems_4)
                        {
                            SqlCommand testCMD1 = new SqlCommand("CreatePurchaseOrderDetails", connection);
                            testCMD1.CommandType = CommandType.StoredProcedure;
                            testCMD1.Parameters.Add(new SqlParameter("@POId", System.Data.SqlDbType.BigInt, 50) { Value = poID });
                            testCMD1.Parameters.Add(new SqlParameter("@Offer_Number", System.Data.SqlDbType.VarChar, 300) { Value = i.Offer_Number.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Vendor_Name", System.Data.SqlDbType.VarChar, 300) { Value = i.Vendor_Name.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Vendor_Code", System.Data.SqlDbType.VarChar, 300) { Value = i.Vendor_Code.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Offer_Date", System.Data.SqlDbType.DateTime, 300) { Value = i.Offer_Date.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Contact_No", System.Data.SqlDbType.NChar, 20) { Value = i.Contact_No.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Contact_Person", System.Data.SqlDbType.VarChar, 300) { Value = i.Contact_Person.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Description", System.Data.SqlDbType.VarChar, 300) { Value = i.Description.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@Quantity", System.Data.SqlDbType.BigInt, 50) { Value = i.Quantity });
                            testCMD1.Parameters.Add(new SqlParameter("@Units", System.Data.SqlDbType.VarChar, 50) { Value = i.Units.Trim() });
                            testCMD1.Parameters.Add(new SqlParameter("@UnitPrice", System.Data.SqlDbType.Decimal, 50) { Value = i.Unit_Price });
                            testCMD1.Parameters.Add(new SqlParameter("@TotalPrice", System.Data.SqlDbType.Decimal, 50) { Value = i.Total_Price });
                            testCMD1.Parameters.Add(new SqlParameter("@HSN", System.Data.SqlDbType.Decimal, 50) { Value = i.HSN });
                            testCMD1.Parameters.Add(new SqlParameter("@GST_Value", System.Data.SqlDbType.Decimal, 50) { Value = i.GST_Value });
                            testCMD1.Parameters.Add(new SqlParameter("@QuoteNo", System.Data.SqlDbType.Int, 50) { Value = 4 });
                            testCMD1.Parameters.Add(new SqlParameter("@CreatedBy", System.Data.SqlDbType.BigInt, 50) { Value = _login.EmployeeID });
                            testCMD1.ExecuteNonQuery(); // read output value from @NewId 
                        }

                        SqlCommand testCMD2 = new SqlCommand("UpdatePurchaseOrderApproval", connection);
                        testCMD2.CommandType = CommandType.StoredProcedure;
                        testCMD2.Parameters.Add(new SqlParameter("@POId", System.Data.SqlDbType.BigInt, 50) { Value = poID });
                        testCMD2.Parameters.Add(new SqlParameter("@ApprovalStatusID", System.Data.SqlDbType.BigInt, 50) { Value = approvalStatus });
                        testCMD2.Parameters.Add(new SqlParameter("@CreatedBy", System.Data.SqlDbType.BigInt, 50) { Value = _login.EmployeeID });
                        testCMD2.ExecuteNonQuery();

                        mailTo = (from a in orderManagementContext.UserMaster
                                  where a.EmployeeId == (from emp in orderManagementContext.Employee where emp.EmployeeId == _login.EmployeeID select emp.ReportsTo).FirstOrDefault()
                                  select a.Email).FirstOrDefault();
                        string body = GeneratePOMail(poID.ToString());

                        isMailSent = SendMail(body, "Purchase Order " + poID + " is Approved by " + _login.UserName + ". Kindly Approve or Deny the Purchase Order ");

                        mailTo = (from a in orderManagementContext.UserMaster
                                  where a.EmployeeId == (from ia in orderManagementContext.IndentApproval where ia.IndentId == indentNo select ia.CreatedBy).FirstOrDefault()
                                  select a.Email).FirstOrDefault();

                        isMailSent = SendMail(body, "Purchase Order " + poID + " is Approved by " + _login.UserName);
                        connection.Close();
                        if (isMailSent)
                        {
                            MessageBox.Show("Purchase Order " + poID + " is approved", "Order Management System", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                        else
                        {
                            MessageBox.Show("Purchase Order " + poID + " is approved but mail is not sent", "Order Management System", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                        DisableFields();
                    }
                }
                this.Cursor = null;
            }
            catch (Exception ex)
            {
                this.Cursor = null;
                MessageBox.Show("An error while approving : " + ex.Message, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void txt_indent_no_LostFocus(object sender, RoutedEventArgs e)
        {
            try
            {
                indentNo = Convert.ToInt64(txt_indent_no.Text);
                var isApproved = (from i in orderManagementContext.IndentApproval where i.IndentId == indentNo && i.ApprovalStatusId == 2 select i).FirstOrDefault();

                if (isApproved != null)
                {
                    var isPOApproved = (from i in orderManagementContext.IndentMaster
                                        join j in orderManagementContext.Pomaster on i.IndentId equals j.IndentId
                                        join k in orderManagementContext.Poapproval on j.PoId equals k.PoId
                                        where i.IndentId == indentNo && k.ApprovalStatusId == 2
                                        select i).FirstOrDefault();
                    if (isPOApproved != null)
                    {
                        var poApprovedID = (from i in orderManagementContext.Pomaster
                                            join j in orderManagementContext.IndentApproval on i.IndentId equals j.IndentId
                                            join k in orderManagementContext.Podetails on i.PoId equals k.PoId
                                            where i.IndentId == indentNo && j.ApprovalStatusId == 2 && k.QNo == 4
                                            select i).FirstOrDefault();
                        //  GetPurchaseOrder(poApprovedID.PoId);
                        FillIndent();
                        EnableFields();
                        // DisableCheckBoxes();
                        // FillApprovalStatusLevel(Convert.ToString(poApprovedID.PoId));
                    }
                    else
                    {
                        //var poID = (from i in orderManagementContext.Pomaster
                        //            join j in orderManagementContext.IndentApproval on i.IndentId equals j.IndentId
                        //            join k in orderManagementContext.Podetails on i.PoId equals k.PoId
                        //            where i.IndentId == indentNo && j.ApprovalStatusId == 2 && k.QNo == 1
                        //            select i).FirstOrDefault();
                        //GetPurchaseOrder(poID.PoId);
                        FillIndent();
                        EnableFields();
                        // DisableCheckBoxes();
                        EnableCheckBoxes();
                        //FillApprovalStatusLevel(Convert.ToString(poID.PoId));
                    }
                }
                else
                {
                    MessageBox.Show("The indent " + indentNo + " is not approved and not eligible to create PO", "Order Management System", MessageBoxButton.OK, MessageBoxImage.Information);
                    DisableFields();
                    DisableCheckBoxes();
                    txt_indent_no.IsEnabled = true;
                    grid_indentdata.ItemsSource = null;
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error occured in fetching Indent Approval Status", "Order Management System", MessageBoxButton.OK, MessageBoxImage.Information);
                grid_indentdata.ItemsSource = null;
            }
        }
        //private static void GeneratePurchaseOrder(Dictionary<string, string> headersAndFooters, List<ExportPO> poData)
        //{
        //    try
        //    {
        //        var workBook = new XLWorkbook();
        //        workBook.AddWorksheet("Report");
        //        var worksheet = workBook.Worksheet("Report");
        private void GeneratePurchaseOrder(Dictionary<string, string> headersAndFooters, ExportPO poData)
        {
            try
            {
                var workBook = new XLWorkbook();
                workBook.AddWorksheet("Report");
                var worksheet = workBook.Worksheet("Report");

                //var imagePath = @"D:\Dinesh\Projects\GitHub\OrderManagementTool\Images\logo.png";
                //var image = worksheet.AddPicture(imagePath)
                //    .MoveTo(worksheet.Cell("A1"))
                //    .Scale(.1).Placement = ClosedXML.Excel.Drawings.XLPicturePlacement.Move;

                var imagePath = @"..\..\..\..\Images\Modified-Image.png";

                if (!File.Exists(imagePath))
                {
                    imagePath = @"C:\Images\Modified-Image.png";

                    if (!File.Exists(imagePath))
                    {
                        imagePath = @"C:\OMS\Images\Modified-Image.png";
                        if (!File.Exists(imagePath))
                        {
                            imagePath = @"..\..\..\Images\Modified-Image.png";

                            if (!File.Exists(imagePath))
                            {
                                imagePath = @"..\..\Images\Modified-Image.png";
                                if (!File.Exists(imagePath))
                                {
                                    imagePath = @"..\Images\Modified-Image.png";

                                    if (!File.Exists(imagePath))
                                    {
                                        imagePath = @"..\Modified-Image.png";

                                        if (!File.Exists(imagePath))
                                        {
                                            imagePath = @"..\..\..\..\..\Images\Modified-Image.png";
                                            if (!File.Exists(imagePath))
                                            {
                                                imagePath = @"..\..\..\..\..\..\Images\Modified-Image.png";
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    string fileContent = string.Empty;
                    try
                    {
                        using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlConnection"].ToString()))
                        {
                            connection.Open();
                            SqlCommand testCMD = new SqlCommand("FetchLookupImage", connection);
                            testCMD.CommandType = CommandType.StoredProcedure;



                            SqlDataReader dataReader = testCMD.ExecuteReader();
                            while (dataReader.Read())
                            {
                                fileContent = dataReader.GetString(0);
                            }
                        }

                    }
                    catch (Exception ex)
                    {

                    }
                    byte[] binaryData = Convert.FromBase64String(fileContent);

                    MemoryStream _stream = new MemoryStream(binaryData);



                    var image = worksheet.AddPicture(_stream)
                        .MoveTo(worksheet.Cell("B1"))
                        .Scale(.075);
                    worksheet.Column("B").Width = 130;
                    //worksheet.Column("G").Width = 130;
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
                    row.Height = 90;
                    var rangeMerged1 = worksheet.Range("A2:G2").Merge();
                    rangeMerged1.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                    rangeMerged1.Style.Border.TopBorder = XLBorderStyleValues.Medium;
                    rangeMerged1.Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                    rangeMerged1.Style.Border.RightBorder = XLBorderStyleValues.Medium;
                    rangeMerged1.Style.Font.Bold = true;
                    rangeMerged1.Style.Font.FontSize = 12;
                    rangeMerged1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    rangeMerged1.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    //rangeMerged1.Style.Fill.BackgroundColor = XLColor.;
                    var rangeMerged2 = worksheet.Range("A3:G3").Merge();
                    rangeMerged2.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                    rangeMerged2.Style.Border.TopBorder = XLBorderStyleValues.Medium;
                    rangeMerged2.Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                    rangeMerged2.Style.Border.RightBorder = XLBorderStyleValues.Medium;
                    rangeMerged2.Style.Font.Bold = true;
                    rangeMerged2.Style.Font.FontSize = 14;
                    rangeMerged2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    rangeMerged2.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    worksheet.Cell("A1").Value = worksheet.Cell("A1").Value + headersAndFooters["CompanyHeader"];
                    rangeMerged1.Value = headersAndFooters["CompanyAddressLine"];
                    rangeMerged2.Value = headersAndFooters["Header"];
                    //worksheet.Cell("A4").Value = headersAndFooters["To"];
                    var rangeTo = worksheet.Range("A4:D9").Merge();
                    rangeTo.Value = headersAndFooters["From"];
                    rangeTo.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    rangeTo.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                    worksheet.Cell("A4").Style.Font.Bold = true;
                    var rangeFrom = worksheet.Range("A10:D14").Merge();
                    rangeFrom.Value = headersAndFooters["To"];
                    rangeFrom.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    rangeFrom.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                    worksheet.Cell("A4").Style.Font.Bold = true;


                    // worksheet.Cell("D4").Value = headersAndFooters["Info"];
                    var rangeInfo = worksheet.Range("E4:G4").Merge();
                    rangeInfo.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    rangeInfo.Value = headersAndFooters["Info"];
                    rangeInfo.Style.Font.Bold = true;

                    worksheet.Cell("B5").Value = null;
                    worksheet.Cell("B6").Value = null;
                    worksheet.Cell("B7").Value = null;
                    worksheet.Cell("B8").Value = null;
                    worksheet.Cell("B9").Value = null;
                    worksheet.Cell("E5").Value = headersAndFooters["PONo"];
                    worksheet.Cell("E6").Value = headersAndFooters["PODate"];
                    worksheet.Cell("E7").Value = headersAndFooters["RefNo"];
                    worksheet.Cell("E8").Value = headersAndFooters["RefDate"];
                    worksheet.Cell("E9").Value = headersAndFooters["Attn"];
                    worksheet.Cell("E5").Style.Font.Bold = true;
                    worksheet.Cell("E6").Style.Font.Bold = true;
                    worksheet.Cell("E7").Style.Font.Bold = true;
                    worksheet.Cell("E8").Style.Font.Bold = true;
                    worksheet.Cell("E9").Style.Font.Bold = true;
                    //worksheet.Cell("A10").Value = headersAndFooters["From"];
                    //worksheet.Cell("A10").Style.Font.Bold = true;
                    var rangeEF1 = worksheet.Range("F5:G5").Merge();
                    worksheet.Cell("F5").Value = poData.POID.ToString().Trim();
                    var rangeEF2 = worksheet.Range("F6:G6").Merge();
                    worksheet.Cell("F6").Value = poData.PODate.ToShortDateString().Trim();
                    var rangeEF3 = worksheet.Range("F7:G7").Merge();
                    worksheet.Cell("F7").Value = poData.Poitems[0].Offer_Number.Trim();
                    var rangeEF4 = worksheet.Range("F8:G8").Merge();
                    worksheet.Cell("F8").Value = poData.Poitems[0].Offer_Date.Split(' ')[0].Trim();
                    var rangeEF5 = worksheet.Range("F9:G9").Merge();
                    worksheet.Cell("F9").Value = poData.Poitems[0].Contact_Person.Trim();

                    worksheet.Cell("F5").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    worksheet.Cell("F6").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    worksheet.Cell("F7").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    worksheet.Cell("F8").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    worksheet.Cell("F9").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    worksheet.Cell("A10").Style.Font.Bold = true;
                    worksheet.Cell("B11").Value = poData.LocatioName;
                    worksheet.Cell("B12").Value = poData.LocationAddressInfo == null ? "" : poData.LocationAddressInfo.Address1;
                    worksheet.Cell("B13").Value = poData.LocationAddressInfo == null ? "" : poData.LocationAddressInfo.Address2;
                    worksheet.Cell("B14").Value = poData.LocationAddressInfo == null ? "" : poData.LocationAddressInfo.Address3;
                    worksheet.Cell("E10").Value = "Mobile No";
                    worksheet.Cell("E11").Value = headersAndFooters["Remarks"];
                    worksheet.Cell("E12").Value = "GST #";
                    worksheet.Cell("E13").Value = "IEC #";
                    worksheet.Cell("E14").Value = "Project Code";
                    worksheet.Cell("E10").Style.Font.Bold = true;
                    worksheet.Cell("E11").Style.Font.Bold = true;
                    worksheet.Cell("E12").Style.Font.Bold = true;
                    worksheet.Cell("E13").Style.Font.Bold = true;
                    worksheet.Cell("E14").Style.Font.Bold = true;
                    // worksheet.Cell("F9").Value = poData.Poitems[0].Contact_No;
                    var rangeEF6 = worksheet.Range("F10:G10").Merge();
                    worksheet.Cell("F10").Value = poData.Poitems[0].Contact_No;
                    worksheet.Cell("F10").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

                    var rangeEF7 = worksheet.Range("F11:G11").Merge();
                    worksheet.Cell("F11").Value = poData.Remarks;
                    var rangeEF8 = worksheet.Range("F12:G12").Merge();
                    worksheet.Cell("F12").Value = headersAndFooters["GSTNo"];
                    var rangeEF9 = worksheet.Range("F13:G13").Merge();
                    worksheet.Cell("F13").Value = headersAndFooters["IECNo"];
                    var rangeEF10 = worksheet.Range("F14:G14").Merge();
                    worksheet.Cell("F14").Value = poData.LocatioName;

                    //worksheet.Cell("F10").Style.Font.Bold = true;
                    worksheet.Cell("F11").Style.Font.Bold = true;
                    worksheet.Cell("F12").Style.Font.Bold = true;
                    worksheet.Cell("F13").Style.Font.Bold = true;
                    worksheet.Cell("F14").Style.Font.Bold = true;
                    // var rangeMerged3 = worksheet.Range("A15:C15").Merge();
                    // worksheet.Cell("A15").Value = headersAndFooters["Materials"];
                    // worksheet.Cell("A15").Style.Font.Bold = true;
                    worksheet.Cell("A15").Value = "S.No";
                    worksheet.Cell("B15").Value = "Description";
                    worksheet.Cell("C15").Value = "Qty";
                    worksheet.Cell("D15").Value = "Units";
                    worksheet.Cell("E15").Value = "Unit Price (INR)";
                    worksheet.Cell("F15").Value = "HSN";
                    worksheet.Cell("G15").Value = "Total Price (INR)";

                    worksheet.Cell("A15").Style.Font.Bold = true;
                    worksheet.Cell("B15").Style.Font.Bold = true;
                    worksheet.Cell("C15").Style.Font.Bold = true;
                    worksheet.Cell("D15").Style.Font.Bold = true;
                    worksheet.Cell("E15").Style.Font.Bold = true;
                    worksheet.Cell("F15").Style.Font.Bold = true;
                    worksheet.Cell("G15").Style.Font.Bold = true;
                    int j = 16;
                    int i = 1;
                    decimal total = 0;
                    foreach (Poitem poInfo in poData.Poitems)
                    {
                        worksheet.Cell("A" + j).Value = i;
                        worksheet.Cell("B" + j).Value = poInfo.Description;
                        worksheet.Cell("B" + j).Style.Alignment.SetWrapText(true);
                        worksheet.Cell("C" + j).Value = poInfo.Quantity;
                        worksheet.Cell("D" + j).Value = poInfo.Units;
                        worksheet.Cell("E" + j).Value = poInfo.Unit_Price;
                        worksheet.Cell("F" + j).Value = poInfo.HSN;
                        worksheet.Cell("G" + j).Value = poInfo.Total_Price;
                        total += poInfo.Total_Price;
                        worksheet.Row(j).Height = (poInfo.Description.Length / 45) * 13;
                        j++;
                        i++;
                    }

                    worksheet.Cell("B" + j).Value = headersAndFooters["SubTotal"];
                    worksheet.Cell("B" + j).Style.Font.Bold = true;
                    worksheet.Cell("G" + j).Value = total;
                    worksheet.Cell("G" + j).Style.Font.Bold = true;

                    j += 1;
                    decimal gst = 0;
                    decimal finalTotal = 0;
                    if (txt_GST.Text != "")
                    {
                        gst = Math.Round(total * (Convert.ToDecimal(txt_GST.Text) / 100));
                        worksheet.Cell("B" + j).Value = headersAndFooters["GST"] + " " + poData.Poitems[0].GST_Value + "%";
                        worksheet.Cell("B" + j).Style.Font.Bold = true;
                        worksheet.Cell("G" + j).Value = Convert.ToString(gst);
                        worksheet.Cell("G" + j).Style.Font.Bold = true;
                        j += 1;
                    }
                    finalTotal = total + gst;
                    worksheet.Cell("B" + j).Value = headersAndFooters["FinalTotal"];
                    worksheet.Cell("B" + j).Style.Font.Bold = true;
                    worksheet.Cell("G" + j).Value = Convert.ToString(finalTotal);
                    worksheet.Cell("G" + j).Style.Font.Bold = true;
                    j += 1;

                    var rangeMerged101 = worksheet.Range("A" + j + ":G" + j).Merge();
                    worksheet.Cell("A" + j).Value = headersAndFooters["Terms"];
                    worksheet.Cell("A" + j).Style.Font.Bold = true;
                    int k1 = j;
                    j += 1;
                    int m = j + 5;
                    // int l = j + 5;
                    //var rangeMerged102 = worksheet.Range("A" + j + ":F" + l).Merge();
                    var rangeMerged1101 = worksheet.Range("A" + j + ":G" + m).Merge();
                    worksheet.Cell("A" + j).Value = headersAndFooters["Terms1"] +"\n" + headersAndFooters["Terms2"]
                        + "\n" + headersAndFooters["Terms3"] + "\n" + headersAndFooters["Terms4"] + "\n" + headersAndFooters["Terms5"];
                    worksheet.Cell("A" + j).Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                    j += 5;
                    //var rangeMerged1102 = worksheet.Range("A" + j + ":G" + j).Merge();
                    //worksheet.Cell("A" + j).Value = headersAndFooters["Terms2"];
                    //j += 1;
                    //var rangeMerged1103 = worksheet.Range("A" + j + ":G" + j).Merge();
                    //worksheet.Cell("A" + j).Value = headersAndFooters["Terms3"];
                    //j += 1;
                    //var rangeMerged1104 = worksheet.Range("A" + j + ":G" + j).Merge();
                    //worksheet.Cell("A" + j).Value = headersAndFooters["Terms4"];
                    //j += 1;
                    //var rangeMerged1105 = worksheet.Range("A" + j + ":G" + j).Merge();
                    //worksheet.Cell("A" + j).Value = headersAndFooters["Terms5"];
                    //+"\n" + headersAndFooters["Terms3"] + "\n" + headersAndFooters["Terms4"] + "\n" + headersAndFooters["Terms5"];
                    //rangeMerged102.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                    int l1 = j;
                    j += 1;
                    var rangeMerged103 = worksheet.Range("A" + j + ":G" + j).Merge();
                    j += 1;
                    var rangeMerged104 = worksheet.Range("A" + j + ":G" + j).Merge();
                    j += 1;
                    var rangeMerged5 = worksheet.Range("A" + j + ":C" + j).Merge();
                    worksheet.Cell("A" + j).Value = headersAndFooters["SpecialInstructions"];
                    worksheet.Cell("A" + j).Style.Font.Bold = true;
                    var rangeMerged6 = worksheet.Range("D" + j + ":G" + j).Merge();
                    worksheet.Cell("E" + j).Value = headersAndFooters["Rejections"];
                    worksheet.Cell("E" + j).Style.Font.Bold = true;
                    j += 1;
                    var rangeMerged7 = worksheet.Range("A" + j + ":C" + j).Merge();
                    worksheet.Cell("A" + j).Value = headersAndFooters["SplLine1"];
                    worksheet.Cell("A" + j).Value = worksheet.Cell("A" + j).Value + "\n" +
                        headersAndFooters["SplLine2"];
                    worksheet.Cell("A" + j).WorksheetRow().Height = 60;
                    rangeMerged7.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                    worksheet.Cell("A" + j).Style.Alignment.WrapText = true;
                    var rangeMerged8 = worksheet.Range("D" + j + ":G" + j).Merge();
                    worksheet.Cell("D" + j).Value = headersAndFooters["RejLine1"];
                    worksheet.Cell("D" + j).Style.Alignment.WrapText = true;
                    int l = j;
                    row = worksheet.Row(1);
                    row.AdjustToContents();
                    row.Height = 40;
                    j += 1;
                    var rangeMerged106 = worksheet.Range("A" + j + ":G" + j).Merge();
                    j += 1;
                    rangeMerged106.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    var rangeMerged107 = worksheet.Range("A" + j + ":G" + j).Merge();
                    rangeMerged107.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    j += 1;
                    var rangeMerged108 = worksheet.Range("A" + j + ":G" + j).Merge();
                    rangeMerged108.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    j += 1;
                    var rangeMerged109 = worksheet.Range("A" + j + ":G" + j).Merge();
                    rangeMerged109.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    int k = j;
                    //var rangeMerged9 = worksheet.Range("A" + k + ":B" + j).Merge();
                    //var rangeMerged10 = worksheet.Range("C" + k + ":D" + j).Merge();
                    //var rangeMerged11 = worksheet.Range("E" + k + ":F" + j).Merge();
                    //worksheet.Cell("A" + j).Value = headersAndFooters["SplLine2"];
                    //worksheet.Cell("A" + j).Style.Alignment.WrapText = true;
                    j += 1;
                    int n = j + 3;
                    var rangeMerged12 = worksheet.Range("A" + j + ":G" + n).Merge();
                    //var rangeMerged13 = worksheet.Range("C" + j + ":D" + j).Merge();
                    //var rangeMerged14 = worksheet.Range("E" + j + ":G" + j).Merge();
                    worksheet.Cell("A" + j).Value = "For SLPP RENEW LLP" +"\n" +"\n"+"\n" +"\n" + headersAndFooters["Sign"]; 
                    worksheet.Cell("A" + j).Style.Font.Bold = true;
                    worksheet.Cell("A" + j).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    rangeMerged12.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    rangeMerged12.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    //j += 1;
                    //var rangeMerged12a = worksheet.Range("A" + j + ":G" + j).Merge();
                    //j += 1;
                    //var rangeMerged12b = worksheet.Range("A" + j + ":G" + j).Merge();
                    //j += 1;
                    //var rangeMerged13 = worksheet.Range("A" + j + ":G" + j).Merge();
                    //worksheet.Cell("A" + j).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    //worksheet.Cell("A" + j).Style.Font.Bold = true;
                    //worksheet.Cell("A" + j).Value = headersAndFooters["Sign"];
                    //worksheet.Cell("C" + j).Style.Font.Bold = true;
                    //worksheet.Cell("C" + j).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    //worksheet.Cell("E" + j).Value = headersAndFooters["ApprovedBy"];
                    //worksheet.Cell("E" + j).Style.Font.Bold = true;
                    //worksheet.Cell("E" + j).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    var rangeRows = worksheet.Range("A3" + ":G" + l);

                    rangeRows.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    rangeRows.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                    rangeRows.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    rangeRows.Style.Border.RightBorder = XLBorderStyleValues.Thin;

                    var rangeRowsNoBorder = worksheet.Range("A"+k1 + ":G" + l1);

                    rangeRowsNoBorder.Style.Border.BottomBorder = XLBorderStyleValues.None;
                    rangeRowsNoBorder.Style.Border.TopBorder = XLBorderStyleValues.None;
                    rangeRowsNoBorder.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                    rangeRowsNoBorder.Style.Border.RightBorder = XLBorderStyleValues.Thin;

                    
                    worksheet.Columns(1, 10).AdjustToContents();
                    worksheet.Column(2).Width = 45;

                   // worksheet.Column("G").Width = 1;

                    worksheet.PageSetup.PaperSize = XLPaperSize.A4Paper;
                    worksheet.PageSetup.PrintAreas.Add("A1:G" + j);

                    //IXLSheetView view = worksheet.SheetView;
                    //view.ZoomScale = 80;

                    worksheet.PageSetup.PageOrder = XLPageOrderValues.OverThenDown;

                    worksheet.PageSetup.Scale = 80;

                    //worksheet.Column(1).Width = 20;
                    if (!Directory.Exists(Convert.ToString(headersAndFooters["ReportGeneratedPath"])))
                    {
                        Directory.CreateDirectory(Convert.ToString(headersAndFooters["ReportGeneratedPath"]));
                    }
                    string filePath = Convert.ToString(headersAndFooters["ReportGeneratedPath"]) +
                        "Purchase_Order_" + poData.POID + "_" + DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() +
                        DateTime.Now.Day.ToString() + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() +
                            DateTime.Now.Second.ToString() + ".xlsx";

                    workBook.SaveAs(filePath);

                    MessageBox.Show("Purchase Order is generated in " + filePath, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show("Error occured while generating report ." + ex.Message, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Information);
                //Console.WriteLine("An Exception occurred. Kindly contact the Administrator");
                //////log.Error("Error while generating for Bot Report");
                //////log.Error(ex.Message);
            }
        }

        private void grid_po_confirmation_Scroll(object sender, System.Windows.Controls.Primitives.ScrollEventArgs e)
        {

        }
    }
}
