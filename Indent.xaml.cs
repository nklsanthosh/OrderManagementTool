using OrderManagementTool.Models;
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
using Caliburn.Micro;
using OrderManagementTool.Models.Indent;
using System.Data;
using OrderManagementTool.Models.LogIn;
using Microsoft.Data.SqlClient;
using System.Configuration;
using OrderManagementTool.Models.Excel;
using ClosedXML.Excel;

namespace OrderManagementTool
{
    /// <summary>
    /// Interaction logic for Indent.xaml
    /// </summary>
    public partial class Indent : Window
    {
        private List<string> itemCode;
        private string selectedItemCode;
        private List<string> units;
        private string selectedItemName;
        private int quantityEntered;
        private int gridSelectedIndex;
        OrderManagementContext orderManagementContext = new OrderManagementContext();
        private List<GridIndent> gridIndents = new List<GridIndent>();
        public BindableCollection<string> ItemName { get; set; }
        private readonly Login _login;

        public Indent(Login login)
        {
            _login = login;
            InitializeComponent();
            LoadItemName();
            txt_raised_by.Text = _login.UserEmail;
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

        private void DataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //DataRowView rowview = grid_indentdata.SelectedItem as DataRowView;
            var rowview = grid_indentdata.SelectedItem as GridIndent;
            gridSelectedIndex = grid_indentdata.SelectedIndex;

            if (rowview != null)
            {
                txt_quantity.Text = Convert.ToInt32(rowview.Quantity).ToString();
                LoadItemName();
                cbx_itemname.SelectedItem = rowview.ItemName;
                LoadItemCode();
                cbx_itemcode.SelectedItem = rowview.ItemCode;
                LoadDescription();
                LoadUnits();
                cbx_units.SelectedItem = rowview.Units;
            }
        }

        private void cbx_itemname_DropDownOpened(object sender, EventArgs e)
        {
            LoadItemName();
        }

        private void cbx_itemcode_DropDownOpened(object sender, EventArgs e)
        {
            LoadItemCode();
        }

        private void cbx_units_DropDownOpened(object sender, EventArgs e)
        {
            LoadUnits();
        }

        private void txt_quantity_lostfocus(object sender, RoutedEventArgs e)
        {
            if (Int32.TryParse(txt_quantity.Text, out int value))
            {
                quantityEntered = Convert.ToInt32(txt_quantity.Text);
            }
            else
            {
                MessageBox.Show("Please enter valid number", "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void cbx_itemcode_DropDownClosed(object sender, EventArgs e)
        {
            LoadDescription();

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (cbx_itemname.SelectedItem != null && cbx_itemcode.SelectedItem != null && txt_quantity.Text != ""
               && quantityEntered != 0 && cbx_units.SelectedItem != null)
            {
                bool itemPresent = false;
                GridIndent gridIndent = new GridIndent();
                gridIndent.SlNo = gridIndents.Count + 1;
                gridIndent.ItemCode = cbx_itemcode.SelectedItem.ToString();
                gridIndent.ItemName = cbx_itemname.SelectedItem.ToString();
                gridIndent.Quantity = quantityEntered;
                gridIndent.Technical_Specifications = txt_technical_description.Text.Trim();
                gridIndent.Units = cbx_units.SelectedItem.ToString();
                gridIndent.Remarks = txt_remarks.Text.Trim();

                foreach (var i in gridIndents)
                {
                    if (i.ItemName == gridIndent.ItemName && 
                            i.ItemCode == gridIndent.ItemCode &&
                                i.Description == gridIndent.Description && 
                                    i.Technical_Specifications == gridIndent.Technical_Specifications && 
                                        i.Units == gridIndent.Units && itemPresent == false )
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
            else if (cbx_itemname.SelectedItem == null)
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
        }

        private void LoadItemName()
        {
            cbx_itemname.Items.Clear();
            var itemName = (from i in orderManagementContext.ItemCategory
                            select i.ItemCategoryName).Distinct().ToList();
            foreach (var i in itemName)
            {
                cbx_itemname.Items.Add(i.Trim());
            }
        }
        private void LoadDescription()
        {
            selectedItemCode = cbx_itemcode.SelectionBoxItem.ToString();
            var itemDetails = (from i in orderManagementContext.ItemMaster
                               from ic in orderManagementContext.ItemCategory
                               where i.ItemCategoryId == ic.ItemCategoryId && 
                               ic.ItemCategoryName == selectedItemName && 
                               i.ItemCode == selectedItemCode
                               select i).FirstOrDefault();
            if (itemDetails != null)
            {
                txt_description.Text = itemDetails.Description;
                txt_technical_description.Text = itemDetails.TechnicalSpecification;
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
            if (cbx_itemname.SelectedItem != null)
            {
                selectedItemName = cbx_itemname.SelectedItem.ToString();
                cbx_itemcode.Items.Clear();
                itemCode = (from i in orderManagementContext.ItemMaster
                            from ic in orderManagementContext.ItemCategory
                            where i.ItemCategoryId == ic.ItemCategoryId && ic.ItemCategoryName == selectedItemName
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
            cbx_itemname.SelectedValue = "";
            cbx_itemcode.SelectedValue = "";
            txt_description.Text = "";
            txt_quantity.Text = "";
            txt_technical_description.Text = "";
            cbx_units.SelectedValue = "";
            txt_remarks.Text = "";
            txt_technical_description.Text = "";
        }

        private void LoadApprovalId()
        {
            // var approval= from i in orderManagementContext.
        }

        private void btn_clear_fields_Click(object sender, RoutedEventArgs e)
        {
            ClearFields();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            if (gridSelectedIndex >= 0)
            {
                // bool itemPresent = false;
                GridIndent value = new GridIndent();
                value.ItemCode = cbx_itemcode.SelectedItem.ToString();
                value.ItemName = cbx_itemname.SelectedItem.ToString();
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
                gridIndent.ItemName = cbx_itemname.SelectedItem.ToString();
                gridIndent.Quantity = quantityEntered;
                gridIndent.Technical_Specifications = txt_technical_description.Text.Trim();
                gridIndent.Units = cbx_units.SelectedItem.ToString();
                gridIndent.Remarks = txt_remarks.Text.Trim();
                grid_indentdata.ItemsSource = null;
                grid_indentdata.ItemsSource = gridIndents;
                gridSelectedIndex = -1;
                ClearFields();
                // }
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
            LoadApprovalId();
        }

        private void btn_create_indent_Click(object sender, RoutedEventArgs e)
        {
            if (datepicker_date1.SelectedDate.ToString() == "")
            {
                MessageBox.Show("Please enter Valid Date");
                return;
            }
            //if (cbx_approval_id.SelectedValue == null)
            //{
            //    MessageBox.Show("Please select Approver");
            //    return;
            //}
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
                // var approvalID= from a in  orderManagementContext.
                // saveIndent.ApprovalID = cbx_approval_id.SelectedItem.ToString();
                saveIndent.ApprovalID = 1;
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


                        foreach (var i in saveIndent.GridIndents)
                        {
                            SqlCommand testCMD1 = new SqlCommand("create_indentDetails", connection);
                            testCMD1.CommandType = CommandType.StoredProcedure;
                            testCMD1.Parameters.Add(new SqlParameter("@IndentID", System.Data.SqlDbType.BigInt, 50) { Value = saveIndent.IndentId });
                            testCMD1.Parameters.Add(new SqlParameter("@ItemName", System.Data.SqlDbType.VarChar, 50) { Value = i.ItemName });
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
                }
                catch(Exception ex)
                {
                    MessageBox.Show("An error occured during save. "+ex.Message, "Order Management System", MessageBoxButton.OK,MessageBoxImage.Error);
                }
            }
        }

        private void btn_generate_report_Click(object sender, RoutedEventArgs e)
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
            GenerateIndent(_headers, headerData, gridIndents);
        }
        private static Dictionary<string, string> GetHeaders()
        {
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

        private static void GenerateIndent(Dictionary<string, string> headersAndFooters,
            IndentHeader headerData, List<GridIndent> gridIndents)
        {
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

                worksheet.Cell("B5").Value = headerData.To;
                worksheet.Cell("B6").Value = headerData.AddressLine1;
                worksheet.Cell("B7").Value = headerData.AddressLine2;
                worksheet.Cell("B8").Value = headerData.AddressLine3;
                worksheet.Cell("B9").Value = headerData.Contact;
                worksheet.Cell("F5").Value = headersAndFooters["IndentNo"];
                worksheet.Cell("F6").Value = headersAndFooters["IDate"];
                worksheet.Cell("F7").Value = headersAndFooters["RefNo"];
                worksheet.Cell("F8").Value = headersAndFooters["RefDate"];
                worksheet.Cell("F9").Value = headersAndFooters["Attn"];
                worksheet.Cell("G5").Value = headerData.IndentNo;
                worksheet.Cell("G6").Value = headerData.IndentDate;
                worksheet.Cell("G7").Value = headerData.RefNo;
                worksheet.Cell("G8").Value = headerData.RefDate;
                worksheet.Cell("G9").Value = headerData.Attention;

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
                worksheet.Cell("G10").Value = headerData.Remarks;
                worksheet.Cell("G11").Value = headerData.GSTNo;
                worksheet.Cell("G12").Value = headerData.IECNo;
                worksheet.Cell("G13").Value = "1 Of 1";


                worksheet.Cell("A16").Value = "S.No";
                worksheet.Cell("B16").Value = "Item Code";
                worksheet.Cell("C16").Value = "Item Name";
                worksheet.Cell("D16").Value = "Description";
                worksheet.Cell("E16").Value = "Units";
                worksheet.Cell("F16").Value = "Quantity";
                worksheet.Cell("G16").Value = "Remarks";
                int j = 17;
                int i = 1;
                decimal total = 0;
                foreach (GridIndent indent in gridIndents)
                {
                    worksheet.Cell("A" + j).Value = i;
                    worksheet.Cell("B" + j).Value = indent.ItemCode;
                    worksheet.Cell("C" + j).Value = indent.ItemName;
                    worksheet.Cell("D" + j).Value = indent.Description;
                    worksheet.Cell("E" + j).Value = indent.Units;
                    worksheet.Cell("F" + j).Value = indent.Quantity;
                    worksheet.Cell("G" + j).Value = indent.Remarks;
                    //total += indent.TotalPrice;
                    j++;
                    i++;
                }

                //worksheet.Cell("B" + j).Value = headersAndFooters["SubTotal"];
                //worksheet.Cell("G" + j).Value = total;

                //j += 1;
                //decimal gst = Math.Round(total * Convert.ToDecimal(.18));
                //decimal finalTotal = total + gst;
                //worksheet.Cell("B" + j).Value = headersAndFooters["GST"];
                //worksheet.Cell("G" + j).Value = Convert.ToString(gst);

                //j += 1;
                //worksheet.Cell("B" + j).Value = headersAndFooters["FinalTotal"];
                //worksheet.Cell("G" + j).Value = Convert.ToString(finalTotal);

                //j += 1;
                //var rangeMerged4 = worksheet.Range("B" + j + ":G" + j).Merge();
                //worksheet.Cell("B" + j).Value = headersAndFooters["Words"];

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
                row.Height = 60;
                j += 1;
                var rangeMerged9 = worksheet.Range("A" + j + ":D" + j).Merge();
                worksheet.Cell("A" + j).Value = headersAndFooters["SplLine2"];
                worksheet.Cell("A" + j).Style.Alignment.WrapText = true;
                j += 3;
                worksheet.Cell("G" + j).Value = headersAndFooters["For"] + " " + headersAndFooters["CompanyHeader"];
                j += 3;
                worksheet.Cell("G" + j).Value = headersAndFooters["Sign"];
                var rangeRows = worksheet.Range("A1" + ":G" + j);

                rangeRows.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                rangeRows.Style.Border.TopBorder = XLBorderStyleValues.Medium;
                rangeRows.Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                rangeRows.Style.Border.RightBorder = XLBorderStyleValues.Medium;
                

                worksheet.Columns(1, 10).AdjustToContents();
                //worksheet.Column(1).Width = 20;
                string filePath = Convert.ToString(headersAndFooters["ReportGeneratedPath"]) +
                    "Indent_" + DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() +
                    DateTime.Now.Day.ToString() + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() +
                        DateTime.Now.Second.ToString() + ".xlsx";

                workBook.SaveAs(filePath);
                MessageBox.Show("The Indent file has been created.", "Order Management System",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                //Console.WriteLine("An Exception occurred. Kindly contact the Administrator");
                //log.Error("Error while generating for Bot Report");
                //log.Error(ex.Message);
                MessageBox.Show("An error occurred. Please contact the administrator", "Order Management System",
                    MessageBoxButton.OK, MessageBoxImage.Error);
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

                //worksheet.Cell("A1").Style.Fill.BackgroundColor = XLColor.Gray;
                //worksheet.Cell("B1").Style.Fill.BackgroundColor = XLColor.Gray;
                //worksheet.Cell("C1").Style.Fill.BackgroundColor = XLColor.Gray;
                //worksheet.Cell("D1").Style.Fill.BackgroundColor = XLColor.Gray;
                //worksheet.Cell("E1").Style.Fill.BackgroundColor = XLColor.Gray;
                //worksheet.Cell("F1").Style.Fill.BackgroundColor = XLColor.Gray;
                //worksheet.Cell("G1").Style.Fill.BackgroundColor = XLColor.Gray;
                //worksheet.Cell("H1").Style.Fill.BackgroundColor = XLColor.Gray;
                //worksheet.Cell("A1").Style.Font.Bold = true;
                //worksheet.Cell("B1").Style.Font.Bold = true;
                //worksheet.Cell("C1").Style.Font.Bold = true;
                //worksheet.Cell("D1").Style.Font.Bold = true;
                //worksheet.Cell("E1").Style.Font.Bold = true;
                //worksheet.Cell("F1").Style.Font.Bold = true;
                //worksheet.Cell("G1").Style.Font.Bold = true;
                //worksheet.Cell("H1").Style.Font.Bold = true;
                //worksheet.Cell("A1").Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("A1").Style.Border.TopBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("A1").Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("A1").Style.Border.RightBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("B1").Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("B1").Style.Border.TopBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("B1").Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("B1").Style.Border.RightBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("C1").Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("C1").Style.Border.TopBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("C1").Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("C1").Style.Border.RightBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("D1").Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("D1").Style.Border.TopBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("D1").Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("D1").Style.Border.RightBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("E1").Style.Border.TopBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("E1").Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("E1").Style.Border.RightBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("E1").Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("F1").Style.Border.TopBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("F1").Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("F1").Style.Border.RightBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("F1").Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("G1").Style.Border.TopBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("G1").Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("G1").Style.Border.RightBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("G1").Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("H1").Style.Border.TopBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("H1").Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("H1").Style.Border.RightBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("H1").Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                //int index = 2;
                //decimal grandTotalAmount = 0;
                //decimal grandTaxTotal = 0;
                //decimal grandBillTotal = 0;
                //foreach (ArasuOutputData _report in lstOutputData)
                //{
                //    worksheet.Cell("A" + index).Value = Convert.ToString(_report.DueDate);
                //    worksheet.Cell("B" + index).Value = Convert.ToString(_report.CustomerId);
                //    worksheet.Cell("C" + index).Value = Convert.ToString(_report.CustomerName);
                //    worksheet.Cell("D" + index).Value = Convert.ToString(_report.SerialNumber);
                //    worksheet.Cell("E" + index).Value = Convert.ToString(_report.CAFNumber);
                //    worksheet.Cell("F" + index).Value = Convert.ToString(_report.BillAmount);
                //    worksheet.Cell("G" + index).Value = Convert.ToString(_report.TaxAmount);
                //    worksheet.Cell("H" + index).Value = Convert.ToString(_report.TotalAmount);
                //    grandTotalAmount += _report.TotalAmount;
                //    grandTaxTotal += _report.TaxAmount;
                //    grandBillTotal += _report.BillAmount;
                //    worksheet.Cell("A" + index).Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("A" + index).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("A" + index).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("A" + index).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("B" + index).Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("B" + index).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("B" + index).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("B" + index).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("C" + index).Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("C" + index).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("C" + index).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("C" + index).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("D" + index).Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("D" + index).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("D" + index).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("D" + index).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("E" + index).Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("E" + index).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("E" + index).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("E" + index).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("F" + index).Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("F" + index).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("F" + index).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("F" + index).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("G" + index).Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("G" + index).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("G" + index).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("G" + index).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("H" + index).Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("H" + index).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("H" + index).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                //    worksheet.Cell("H" + index).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                //    index++;
                //}
                //worksheet.Cell("E" + index).Value = "Grand Total";
                //worksheet.Cell("E" + index).Style.Font.Bold = true;
                //worksheet.Cell("F" + index).Value = grandBillTotal;
                //worksheet.Cell("G" + index).Value = grandTaxTotal;
                //worksheet.Cell("H" + index).Value = grandTotalAmount;
                //worksheet.Cell("E" + index).Style.Font.Bold = true;
                //worksheet.Cell("F" + index).Style.Font.Bold = true;
                //worksheet.Cell("G" + index).Style.Font.Bold = true;
                //worksheet.Cell("H" + index).Style.Font.Bold = true;
                //worksheet.Cell("E" + index).Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("E" + index).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("E" + index).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("E" + index).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("F" + index).Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("F" + index).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("F" + index).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("F" + index).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("G" + index).Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("G" + index).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("G" + index).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("G" + index).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("H" + index).Style.Border.BottomBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("H" + index).Style.Border.TopBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("H" + index).Style.Border.LeftBorder = XLBorderStyleValues.Medium;
                //worksheet.Cell("H" + index).Style.Border.RightBorder = XLBorderStyleValues.Medium;

                worksheet.Columns(1, 10).AdjustToContents();
                //worksheet.Column(1).Width = 20;
                string filePath = Convert.ToString(headersAndFooters["ReportGeneratedPath"]) +
                    "Indent_" + DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() +
                    DateTime.Now.Day.ToString() + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() +
                        DateTime.Now.Second.ToString() + ".xlsx";

                workBook.SaveAs(filePath);
                //outputFilePath = filePath;
                ////Console.WriteLine("Report Generated..");
                //Console.WriteLine("Press any key to continue..");
                //Console.ReadLine();
            }
            catch (Exception ex)
            {
                //Console.WriteLine("An Exception occurred. Kindly contact the Administrator");
                //log.Error("Error while generating for Bot Report");
                //log.Error(ex.Message);
            }
        }
    }
}
