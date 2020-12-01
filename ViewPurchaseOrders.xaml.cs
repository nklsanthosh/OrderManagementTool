using Microsoft.Data.SqlClient;
using OrderManagementTool.Models;
using OrderManagementTool.Models.LogIn;
using OrderManagementTool.Models.Purchase_Order;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace OrderManagementTool
{
    /// <summary>
    /// Interaction logic for ViewPurchaseOrders.xaml
    /// </summary>
    public partial class ViewPurchaseOrders : Window
    {
        OrderManagementContext orderManagementContext = new OrderManagementContext();
        private readonly Login _login;
        private bool isWindowNormallyClosed = true;
        private static int gridSelectedIndex = -1;
        private static List<POView> viewPOs = new List<POView>();
        private long selectedIndentID;

        public ViewPurchaseOrders(Login login)
        {
            _login = login;
            InitializeComponent();
            LoadApprovalStatus();
            LoadGrid();
        }

        private void LoadApprovalStatus()
        {
            try
            {
                cbx_approval_status.SelectedValuePath = "Key";
                cbx_approval_status.DisplayMemberPath = "Value";
                cbx_approval_status.Items.Clear();
                var itemCategoryName = (from i in orderManagementContext.ApprovalStatus
                                        select new
                                        {
                                            i.ApprovalStatusId,
                                            i.ApprovalStatus1
                                        }).Distinct().ToList();
                foreach (var i in itemCategoryName)
                {
                    cbx_approval_status.Items.Add(new KeyValuePair<long, string>(i.ApprovalStatusId, i.ApprovalStatus1.Trim()));
                }
                cbx_approval_status.SelectedValue = 1;
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured during Approval Status Load. " + ex.Message, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void cbx_approval_status_DropDownOpened(object sender, EventArgs e)
        {
            this.Cursor = Cursors.Wait;

        }
        private void cbx_approval_status_DropDownClosed(object sender, EventArgs e)
        {
            LoadGrid();
            this.Cursor = null;
        }

        private void grid_all_POs_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //DataRowView rowview = grid_all_POs.SelectedItem as DataRowView;
            try
            {
                var rowview = grid_all_POs.SelectedItem as POView;
                gridSelectedIndex = grid_all_POs.SelectedIndex;
                if (rowview != null)
                {
                    selectedIndentID = rowview.PO_ID;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred :" + ex.Message,
                                   "Order Management System",
                                       MessageBoxButton.OK,
                                           MessageBoxImage.Error);
                //log.Error("Error on clicking the indent row : " + ex.StackTrace);
            }
        }

        private void btn_view_PO_Click(object sender, RoutedEventArgs e)
        {
            if (gridSelectedIndex >= 0 && selectedIndentID > 0)
            {
                var idFound = (from i in orderManagementContext.Poapproval
                               where i.PoId == selectedIndentID
                               select i.PoId).FirstOrDefault();

                if (idFound != null)
                {
                    QuoteComparer qc = new QuoteComparer(_login, idFound);
                    qc.Show();
                    isWindowNormallyClosed = false;
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Please enter valid PO ID",
                              "Order Management System",
                                  MessageBoxButton.OK,
                                      MessageBoxImage.Error);
                }
            }
            else
            {
                MessageBox.Show("Please select a item in the grid",
                                 "Order Management System",
                                     MessageBoxButton.OK,
                                         MessageBoxImage.Error);
            }
        }

        private void LoadGrid()
        {
            //log.Info("LKoading ...");
            try
            {
                long approvalStatusId = Convert.ToInt64(cbx_approval_status.SelectedValue);
                using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlConnection"].ToString()))
                {
                    connection.Open();
                    SqlCommand testCMD = new SqlCommand("GetPurchaseOrderByApprovalStatus", connection);
                    testCMD.CommandType = CommandType.StoredProcedure;

                    //testCMD.Parameters.Add(new SqlParameter("@UserId", System.Data.SqlDbType.BigInt, 50) { Value = _login.EmployeeID });
                    testCMD.Parameters.Add(new SqlParameter("@UserID", System.Data.SqlDbType.VarChar, 300) { Value = _login.EmployeeID });
                    testCMD.Parameters.Add(new SqlParameter("@ApprovalStatusId", System.Data.SqlDbType.BigInt, 50) { Value = approvalStatusId });

                    // SqlDataReader dataReader = testCMD.ExecuteReader();
                    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(testCMD);

                    DataSet dataSet = new DataSet();
                    sqlDataAdapter.Fill(dataSet);

                    int counter = 0;

                    if (viewPOs.Count > 0)
                    {
                        viewPOs.Clear();
                    }

                    while (counter < dataSet.Tables[0].Rows.Count)
                    {
                        POView viewPO = new POView();
                        viewPO.Sl_No = counter + 1;
                        viewPO.PO_ID = Convert.ToInt64(dataSet.Tables[0].Rows[counter]["PO_ID"]);
                        viewPO.Indent_No = Convert.ToInt64(dataSet.Tables[0].Rows[counter]["IndentID"]);
                        viewPO.Approval_Status_Id = Convert.ToInt64(dataSet.Tables[0].Rows[counter]["ApprovalStatusId"]);
                        viewPO.Approver_Id = Convert.ToInt64(dataSet.Tables[0].Rows[counter]["ApproverId"]);
                        viewPO.Offer_Number = Convert.ToString(dataSet.Tables[0].Rows[counter]["Offer_Number"]);
                        viewPO.Description = Convert.ToString(dataSet.Tables[0].Rows[counter]["Description"]);
                        viewPO.Vendor_Name = Convert.ToString(dataSet.Tables[0].Rows[counter]["Vendor_Name"]);
                        viewPO.Vendor_Code = Convert.ToString(dataSet.Tables[0].Rows[counter]["Vendor_Code"]);
                        viewPO.Offer_Date = Convert.ToString(dataSet.Tables[0].Rows[counter]["Offer_Date"]);
                        viewPO.Contact_No = Convert.ToString(dataSet.Tables[0].Rows[counter]["Contact_Number"]);
                        viewPO.Contact_Person = Convert.ToString(dataSet.Tables[0].Rows[counter]["Contact_Person"]);
                        viewPO.GST_Value = Convert.ToDecimal(dataSet.Tables[0].Rows[counter]["GST_Value"]);
                        viewPO.Quantity = Convert.ToInt32(dataSet.Tables[0].Rows[counter]["Quantity"]);
                        viewPO.Q_No = Convert.ToInt32(dataSet.Tables[0].Rows[counter]["Q_No"]);
                        viewPO.Units = Convert.ToString(dataSet.Tables[0].Rows[counter]["Units"]);
                        viewPO.Unit_Price = Convert.ToInt64(dataSet.Tables[0].Rows[counter]["Unit_Price"]);
                        viewPO.Total_Price = Convert.ToDecimal(dataSet.Tables[0].Rows[counter]["Total_Price"]);

                        viewPOs.Add(viewPO);
                        counter++;
                    }
                    dataSet.Dispose();
                    grid_all_POs.ItemsSource = null;
                    grid_all_POs.ItemsSource = viewPOs;
                }
                //log.Info("Indent Loaded...");
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured during indent load. " + ex.Message, "Order Management System", MessageBoxButton.OK, MessageBoxImage.Error);

                this.Close();
                //log.Error("Eror loading indent : " + ex.StackTrace);
            }
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            if (isWindowNormallyClosed)
            {
                Menu menu = new Menu(_login);
                menu.Show();
            }
        }
    }
}
