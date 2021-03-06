﻿//using log4net;
using Microsoft.Data.SqlClient;
using OrderManagementTool.Models;
using OrderManagementTool.Models.Indent;
using OrderManagementTool.Models.LogIn;
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
    /// Interaction logic for ViewIndent.xaml
    /// </summary>
    public partial class ViewIndents : Window
    {
        //ILog log = LogManager.GetLogger(typeof(MainWindow));
        OrderManagementContext orderManagementContext = new OrderManagementContext();
        private static List<ViewIndent> viewIndents = new List<ViewIndent>();
        private static int gridSelectedIndex = -1;
        private static long selectedIndentID;
        private readonly Login _login;
        private bool isWindowNormallyClosed = true;

        public ViewIndents(Login login)
        {
            _login = login;
            //log.Info("In View Indent...");
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

        private void LoadGrid()
        {
            //log.Info("LKoading ...");
            try
            {
                long approvalStatusId = Convert.ToInt64(cbx_approval_status.SelectedValue);
                using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlConnection"].ToString()))
                {
                    connection.Open();
                    SqlCommand testCMD = new SqlCommand("GetIndentByApprovalStatus", connection);
                    testCMD.CommandType = CommandType.StoredProcedure;

                    //testCMD.Parameters.Add(new SqlParameter("@UserId", System.Data.SqlDbType.BigInt, 50) { Value = _login.EmployeeID });
                    testCMD.Parameters.Add(new SqlParameter("@UserID", System.Data.SqlDbType.VarChar, 300) { Value = _login.EmployeeID });
                    testCMD.Parameters.Add(new SqlParameter("@ApprovalStatusId", System.Data.SqlDbType.BigInt, 50) { Value = approvalStatusId });

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
                        viewIndent.Remarks = Convert.ToString(dataSet.Tables[0].Rows[counter]["Item Remarks"]);

                        viewIndent.Email = Convert.ToString(dataSet.Tables[0].Rows[counter]["Email"]);

                        viewIndents.Add(viewIndent);
                        counter++;
                    }
                    dataSet.Dispose();
                    grid_all_indents.ItemsSource = null;
                    grid_all_indents.ItemsSource = viewIndents;
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

        private void btn_view_indent_Click(object sender, RoutedEventArgs e)
        {
            if (gridSelectedIndex >= 0 && selectedIndentID > 0)
            {
                var idFound = (from i in orderManagementContext.IndentApproval
                               where i.IndentId == selectedIndentID
                               select i.IndentId).FirstOrDefault();

                if (idFound != null)
                {
                    Indent indent = new Indent(_login, idFound);
                    indent.Show();
                    isWindowNormallyClosed = false;
                    this.Close();
                }
                else
                {
                    MessageBox.Show("Please enter valid IndentID",
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

        protected override void OnClosing(CancelEventArgs e)
        {
            if (isWindowNormallyClosed)
            {
                Menu menu = new Menu(_login);
                menu.Show();
            }
        }

        private void grid_all_indents_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //DataRowView rowview = grid_indentdata.SelectedItem as DataRowView;
            try
            {
                var rowview = grid_all_indents.SelectedItem as ViewIndent;
                gridSelectedIndex = grid_all_indents.SelectedIndex;
                if (rowview != null)
                {
                    selectedIndentID = rowview.IndentId;
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
    }
}
