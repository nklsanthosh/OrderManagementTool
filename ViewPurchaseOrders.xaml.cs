using OrderManagementTool.Models;
using OrderManagementTool.Models.LogIn;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace OrderManagementTool
{
    /// <summary>
    /// Interaction logic for ViewPurchaseOrders.xaml
    /// </summary>
    public partial class ViewPurchaseOrders : Window
    {
        OrderManagementContext orderManagementContext = new OrderManagementContext();
        private readonly Login _login;

        public ViewPurchaseOrders(Login login)
        {
            _login = login;
            InitializeComponent();
            LoadApprovalStatus();
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

        private void grid_all_indents_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //DataRowView rowview = grid_indentdata.SelectedItem as DataRowView;
            //try
            //{
            //    var rowview = grid_all_indents.SelectedItem as ViewIndent;
            //    gridSelectedIndex = grid_all_indents.SelectedIndex;
            //    if (rowview != null)
            //    {
            //        selectedIndentID = rowview.IndentId;
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("An error occurred :" + ex.Message,
            //                       "Order Management System",
            //                           MessageBoxButton.OK,
            //                               MessageBoxImage.Error);
            //    //log.Error("Error on clicking the indent row : " + ex.StackTrace);
            //}
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            Menu menu = new Menu(_login);
            menu.Show();
        }
    }
}
