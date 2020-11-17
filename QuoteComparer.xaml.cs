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
        private long poID;

        public QuoteComparer(Login login)
        {
            _login = login;
            InitializeComponent();
            txt_indent_no.IsReadOnly = false;
        }

        public QuoteComparer(Login login, long indentNo)
        {
            _login = login;
            InitializeComponent();
            txt_indent_no.Text = indentNo.ToString();
            FillIndent();
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
                        viewIndent.IndentId = (long)Convert.ToInt64(dataSet.Tables[0].Rows[counter]["IndentID"]);
                        viewIndent.ApproverName = Convert.ToString(dataSet.Tables[0].Rows[counter]["Approver"]);
                        viewIndent.Approval_Status = Convert.ToString(dataSet.Tables[0].Rows[counter]["ApprovalStatus"]);
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

        private void DisableFields()
        {
            txt_indent_no.IsEnabled = false;
            checkbox_Approve1.IsEnabled = false;
            checkbox_Approve2.IsEnabled = false;
            checkbox_Approve3.IsEnabled = false;
            btn_upload.IsEnabled = false;
            btn_upload_2.IsEnabled = false;
            btn_upload_3.IsEnabled = false;
            txt_PO_Remarks.IsEnabled = false;
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

        private void txt_PO_Remarks_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void btn_Create_PO_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                bool isMailSent = false;
                if (poitems_1.Count == 0 && poitems_2.Count == 0 && poitems_3.Count == 0)
                {
                    MessageBox.Show("Please upload any one Quotation",
                          "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                if (txt_indent_no.Text == null || txt_indent_no.Text == "")
                {
                    MessageBox.Show("Please enter indent number",
                         "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
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
                            testCMD1.Parameters.Add(new SqlParameter("@Description", System.Data.SqlDbType.VarChar, 300) { Value = i.Description });
                            testCMD1.Parameters.Add(new SqlParameter("@Quantity", System.Data.SqlDbType.BigInt, 50) { Value = i.Quantity });
                            testCMD1.Parameters.Add(new SqlParameter("@Units", System.Data.SqlDbType.VarChar, 50) { Value = i.Units });
                            testCMD1.Parameters.Add(new SqlParameter("@UnitPrice", System.Data.SqlDbType.Decimal, 50) { Value = i.Unit_Price });
                            testCMD1.Parameters.Add(new SqlParameter("@TotalPrice", System.Data.SqlDbType.Decimal, 50) { Value = i.Total_Price });
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
                            testCMD1.Parameters.Add(new SqlParameter("@Description", System.Data.SqlDbType.VarChar, 300) { Value = i.Description });
                            testCMD1.Parameters.Add(new SqlParameter("@Quantity", System.Data.SqlDbType.BigInt, 50) { Value = i.Quantity });
                            testCMD1.Parameters.Add(new SqlParameter("@Units", System.Data.SqlDbType.VarChar, 50) { Value = i.Units });
                            testCMD1.Parameters.Add(new SqlParameter("@UnitPrice", System.Data.SqlDbType.Decimal, 50) { Value = i.Unit_Price });
                            testCMD1.Parameters.Add(new SqlParameter("@TotalPrice", System.Data.SqlDbType.Decimal, 50) { Value = i.Total_Price });
                            testCMD1.Parameters.Add(new SqlParameter("@QuoteNo", System.Data.SqlDbType.Int, 50) { Value = 2 });
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
                            testCMD1.Parameters.Add(new SqlParameter("@Description", System.Data.SqlDbType.VarChar, 300) { Value = i.Description });
                            testCMD1.Parameters.Add(new SqlParameter("@Quantity", System.Data.SqlDbType.BigInt, 50) { Value = i.Quantity });
                            testCMD1.Parameters.Add(new SqlParameter("@Units", System.Data.SqlDbType.VarChar, 50) { Value = i.Units });
                            testCMD1.Parameters.Add(new SqlParameter("@UnitPrice", System.Data.SqlDbType.Decimal, 50) { Value = i.Unit_Price });
                            testCMD1.Parameters.Add(new SqlParameter("@TotalPrice", System.Data.SqlDbType.Decimal, 50) { Value = i.Total_Price });
                            testCMD1.Parameters.Add(new SqlParameter("@QuoteNo", System.Data.SqlDbType.Int, 50) { Value = 3 });
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

                    //string body = GenerateIndent(indentNo.ToString());
                    isMailSent = SendMail("Purchase Order " + poID + " is created by " + _login.UserName);
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
                    DisableFields();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error Occured during PO Creation." + ex.Message,
                      "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private bool SendMail(string subject)
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

                    mm.Body = subject + " Kindly Approve or Deny the Purchase Order ";
                    mm.IsBodyHtml = true;

                    // mm.Attachments.Add(new System.Net.Mail.Attachment(filePathLocation));

                    SmtpClient smtp = new SmtpClient();
                    smtp.Host = ConfigurationManager.AppSettings["Host"];
                    smtp.EnableSsl = false;
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
        //private static void GeneratePurchaseOrder(Dictionary<string, string> headersAndFooters, List<ExportPO> poData)
        //{
        //    try
        //    {
        //        var workBook = new XLWorkbook();
        //        workBook.AddWorksheet("Report");
        //        var worksheet = workBook.Worksheet("Report");

        //        var imagePath = @"D:\Dinesh\Projects\GitHub\OrderManagementTool\Images\logo.png";
        //        var image = worksheet.AddPicture(imagePath)
        //            .MoveTo(worksheet.Cell("A1"))
        //            .Scale(.1).Placement = ClosedXML.Excel.Drawings.XLPicturePlacement.Move;
        //        var rangeMerged = worksheet.Range("A1:G1").Merge();
        //        rangeMerged.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
        //        rangeMerged.Style.Border.TopBorder = XLBorderStyleValues.Medium;
        //        rangeMerged.Style.Border.LeftBorder = XLBorderStyleValues.Medium;
        //        rangeMerged.Style.Border.RightBorder = XLBorderStyleValues.Medium;
        //        rangeMerged.Style.Font.Bold = true;
        //        rangeMerged.Style.Font.FontSize = 16;
        //        rangeMerged.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        //        rangeMerged.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        //        var row = worksheet.Row(1);
        //        row.Height = 40;
        //        var rangeMerged1 = worksheet.Range("A2:G2").Merge();
        //        rangeMerged1.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
        //        rangeMerged1.Style.Border.TopBorder = XLBorderStyleValues.Medium;
        //        rangeMerged1.Style.Border.LeftBorder = XLBorderStyleValues.Medium;
        //        rangeMerged1.Style.Border.RightBorder = XLBorderStyleValues.Medium;
        //        rangeMerged1.Style.Font.Bold = true;
        //        rangeMerged1.Style.Font.FontSize = 16;
        //        rangeMerged1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        //        rangeMerged1.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        //        var rangeMerged2 = worksheet.Range("A3:G3").Merge();
        //        rangeMerged2.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
        //        rangeMerged2.Style.Border.TopBorder = XLBorderStyleValues.Medium;
        //        rangeMerged2.Style.Border.LeftBorder = XLBorderStyleValues.Medium;
        //        rangeMerged2.Style.Border.RightBorder = XLBorderStyleValues.Medium;
        //        rangeMerged2.Style.Font.Bold = true;
        //        rangeMerged2.Style.Font.FontSize = 16;
        //        rangeMerged2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        //        rangeMerged2.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        //        worksheet.Cell("A1").Value = worksheet.Cell("A1").Value + headersAndFooters["CompanyHeader"];
        //        worksheet.Cell("A2").Value = headersAndFooters["CompanyAddressLine"];
        //        worksheet.Cell("A3").Value = headersAndFooters["Header"];
        //        worksheet.Cell("A4").Value = headersAndFooters["To"];
        //        worksheet.Cell("D4").Value = headersAndFooters["Info"];

        //        worksheet.Cell("B5").Value = "M/S GILIYAL INDUSTRIES";
        //        worksheet.Cell("B6").Value = "Address Line 1";
        //        worksheet.Cell("B7").Value = "Address Line 2";
        //        worksheet.Cell("B8").Value = "Address Line 3";
        //        worksheet.Cell("B9").Value = "Address Line 4";
        //        worksheet.Cell("F5").Value = headersAndFooters["PONo"];
        //        worksheet.Cell("F6").Value = headersAndFooters["PODate"];
        //        worksheet.Cell("F7").Value = headersAndFooters["RefNo"];
        //        worksheet.Cell("F8").Value = headersAndFooters["RefDate"];
        //        worksheet.Cell("F9").Value = headersAndFooters["Attn"];
        //        worksheet.Cell("G5").Value = "PO Number";
        //        worksheet.Cell("G6").Value = "PO Date";
        //        worksheet.Cell("G7").Value = "Ref No";
        //        worksheet.Cell("G8").Value = "Ref Date";
        //        worksheet.Cell("G9").Value = "Attention";

        //        var rangeMerged3 = worksheet.Range("A10:C10").Merge();
        //        worksheet.Cell("A10").Value = headersAndFooters["Materials"];

        //        worksheet.Cell("B11").Value = headersAndFooters["CompanyName"];
        //        worksheet.Cell("B12").Value = headersAndFooters["CompanyAddressLine1"];
        //        worksheet.Cell("B13").Value = headersAndFooters["CompanyAddressLine2"];
        //        worksheet.Cell("B14").Value = headersAndFooters["CompanyContactNo"];
        //        worksheet.Cell("F10").Value = headersAndFooters["Remarks"];
        //        worksheet.Cell("F11").Value = headersAndFooters["GSTNo"];
        //        worksheet.Cell("F12").Value = headersAndFooters["IECNo"];
        //        worksheet.Cell("F13").Value = headersAndFooters["Page"];
        //        // worksheet.Cell("F9").Value = headersAndFooters["Attn"];
        //        worksheet.Cell("G10").Value = "Remarks";
        //        worksheet.Cell("G11").Value = "GST #";
        //        worksheet.Cell("G12").Value = "IEC #";
        //        worksheet.Cell("G13").Value = "1 Of 1";
        //        worksheet.Cell("G14").Value = "Attention";

        //        worksheet.Cell("A16").Value = "S.No";
        //        worksheet.Cell("B16").Value = "Description";
        //        worksheet.Cell("C16").Value = "Qty";
        //        worksheet.Cell("D16").Value = "Units";
        //        worksheet.Cell("E16").Value = "Unit Price (INR)";
        //        worksheet.Cell("F16").Value = "HSN";
        //        worksheet.Cell("G16").Value = "Total Price (INR)";
        //        int j = 17;
        //        int i = 1;
        //        decimal total = 0;
        //        foreach (ExportIndent indent in gridIndents)
        //        {
        //            worksheet.Cell("A" + j).Value = i;
        //            worksheet.Cell("B" + j).Value = indent.Description;
        //            worksheet.Cell("C" + j).Value = indent.Quantity;
        //            worksheet.Cell("D" + j).Value = indent.Units;
        //            worksheet.Cell("E" + j).Value = indent.UnitPrice;
        //            worksheet.Cell("F" + j).Value = indent.HSN;
        //            worksheet.Cell("G" + j).Value = indent.TotalPrice;
        //            total += indent.TotalPrice;
        //            j++;
        //            i++;
        //        }

        //        worksheet.Cell("B" + j).Value = headersAndFooters["SubTotal"];
        //        worksheet.Cell("G" + j).Value = total;

        //        j += 1;
        //        decimal gst = Math.Round(total * Convert.ToDecimal(.18));
        //        decimal finalTotal = total + gst;
        //        worksheet.Cell("B" + j).Value = headersAndFooters["GST"];
        //        worksheet.Cell("G" + j).Value = Convert.ToString(gst);

        //        j += 1;
        //        worksheet.Cell("B" + j).Value = headersAndFooters["FinalTotal"];
        //        worksheet.Cell("G" + j).Value = Convert.ToString(finalTotal);

        //        j += 1;
        //        var rangeMerged4 = worksheet.Range("B" + j + ":G" + j).Merge();
        //        worksheet.Cell("B" + j).Value = headersAndFooters["Words"];

        //        j += 2;
        //        var rangeMerged101 = worksheet.Range("A" + j + ":G" + j).Merge();
        //        worksheet.Cell("A" + j).Value = headersAndFooters["Terms"];
        //        j += 1;
        //        var rangeMerged102 = worksheet.Range("A" + j + ":G" + j).Merge();
        //        worksheet.Cell("A" + j).Value = headersAndFooters["Terms1"];

        //        j += 1;
        //        var rangeMerged103 = worksheet.Range("A" + j + ":G" + j).Merge();
        //        worksheet.Cell("A" + j).Value = headersAndFooters["Terms2"];
        //        j += 1;
        //        var rangeMerged104 = worksheet.Range("A" + j + ":G" + j).Merge();
        //        worksheet.Cell("A" + j).Value = headersAndFooters["Terms3"];
        //        j += 1;
        //        var rangeMerged105 = worksheet.Range("A" + j + ":G" + j).Merge();
        //        worksheet.Cell("A" + j).Value = headersAndFooters["Terms4"];
        //        j += 1;
        //        var rangeMerged106 = worksheet.Range("A" + j + ":G" + j).Merge();
        //        worksheet.Cell("A" + j).Value = headersAndFooters["Terms5"];
        //        j += 3;
        //        var rangeMerged5 = worksheet.Range("A" + j + ":D" + j).Merge();
        //        worksheet.Cell("A" + j).Value = headersAndFooters["SpecialInstructions"];
        //        var rangeMerged6 = worksheet.Range("E" + j + ":G" + j).Merge();
        //        worksheet.Cell("E" + j).Value = headersAndFooters["Rejections"];
        //        j += 1;
        //        var rangeMerged7 = worksheet.Range("A" + j + ":D" + j).Merge();
        //        worksheet.Cell("A" + j).Value = headersAndFooters["SplLine1"];
        //        worksheet.Cell("A" + j).Style.Alignment.WrapText = true;
        //        var rangeMerged8 = worksheet.Range("E" + j + ":G" + j).Merge();
        //        worksheet.Cell("E" + j).Value = headersAndFooters["RejLine1"];
        //        worksheet.Cell("E" + j).Style.Alignment.WrapText = true;
        //        row = worksheet.Row(1);
        //        row.Height = 40;
        //        j += 1;
        //        var rangeMerged9 = worksheet.Range("A" + j + ":D" + j).Merge();
        //        worksheet.Cell("A" + j).Value = headersAndFooters["SplLine2"];
        //        worksheet.Cell("A" + j).Style.Alignment.WrapText = true;
        //        j += 3;
        //        worksheet.Cell("G" + j).Value = headersAndFooters["For"] + " " + headersAndFooters["CompanyHeader"];
        //        j += 3;
        //        worksheet.Cell("G" + j).Value = headersAndFooters["Sign"];

        //        worksheet.Columns(1, 10).AdjustToContents();
        //        //worksheet.Column(1).Width = 20;
        //        string filePath = Convert.ToString(headersAndFooters["ReportGeneratedPath"]) +
        //            "Indent_" + DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() +
        //            DateTime.Now.Day.ToString() + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() +
        //                DateTime.Now.Second.ToString() + ".xlsx";

        //        workBook.SaveAs(filePath);
        //    }
        //    catch (Exception ex)
        //    {
        //        //Console.WriteLine("An Exception occurred. Kindly contact the Administrator");
        //        //////log.Error("Error while generating for Bot Report");
        //        //////log.Error(ex.Message);
        //    }
        //}

    }
}
