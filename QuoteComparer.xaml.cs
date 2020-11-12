using Microsoft.Data.SqlClient;
using OrderManagementTool.Models;
using OrderManagementTool.Models.Indent;
using OrderManagementTool.Models.LogIn;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
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


        public QuoteComparer(Login login)
        {
            _login = login;
            InitializeComponent();
        }

        private void grid_po_confirmation_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void txt_indent_no_LostFocus(object sender, RoutedEventArgs e)
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
    }
}
