//using log4net;
using Microsoft.Data.SqlClient;
using OrderManagementTool.Models;
using OrderManagementTool.Models.LogIn;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Windows;

namespace OrderManagementTool
{
    /// <summary>
    /// Interaction logic for Menu.xaml
    /// </summary>
    public partial class Menu : Window
    {
        //ILog log = LogManager.GetLogger(typeof(MainWindow));
        private readonly Login _login;
        OrderManagementContext orderManagementContext = new OrderManagementContext();
        public Menu()
        {
            //log.Info("In Menu Landing Page...");
            InitializeComponent();
        }

        public Menu(Login login)
        {
            _login = login;
            InitializeComponent();
            try
            {
                //log.Info("Fetching Roles...");

                var rolesMapping = (from roleMap in orderManagementContext.RoleMapping
                                    join roles in orderManagementContext.RolesMaster on roleMap.RoleId equals roles.RoleId
                                    join users in orderManagementContext.UserMaster on roleMap.UserId equals users.UserId
                                    where users.Email.Equals(_login.UserEmail)
                                    orderby roles.RoleId
                                    select new
                                    {
                                        roleMap.RoleId
                                    });
                if (rolesMapping != null)
                {
                    if (rolesMapping.FirstOrDefault().RoleId != 1)
                    {
                        lbl_item.Visibility = Visibility.Hidden;
                        btn_add_item.Visibility = Visibility.Hidden;
                    }
                    else
                    {
                        lbl_item.Visibility = Visibility.Visible;
                        btn_add_item.Visibility = Visibility.Visible;
                    }
                }
                //log.Info("Roles applied..");
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured while fetching user information. " + ex.Message,
                                    "Order Management System",
                                        MessageBoxButton.OK,
                                            MessageBoxImage.Error);
                //log.Error("Role mapping issue : " + ex.StackTrace);
            }
        }


        private void btn_create_Indent_Click(object sender, RoutedEventArgs e)
        {
            Indent indent = new Indent(_login);
            indent.Show();
            this.Close();
        }

        private void btn_add_item_Click(object sender, RoutedEventArgs e)
        {
            Item item = new Item(_login);
            item.Show();
            this.Close();
        }

        private void btn_log_off_Click(object sender, RoutedEventArgs e)
        {
            MainWindow mainWindow = new MainWindow();
            mainWindow.Show();
            this.Close();
        }

        private void btn_create_PO_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                QuoteComparer qC = new QuoteComparer(_login);
                qC.Show();
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured PO Access Check " + ex.Message,
                                   "Order Management System",
                                       MessageBoxButton.OK,
                                           MessageBoxImage.Error);
            }

        }

        private void btn_search_PO_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (txt_po_indent_no.Text != "")
                {
                    long POID = Convert.ToInt64(txt_po_indent_no.Text);

                    // var isFound = (from i in orderManagementContext.Poapproval where i.PoId == POID select i).FirstOrDefault();

                    int? isFound = null;

                    using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlConnection"].ToString()))
                    {
                        connection.Open();

                        SqlCommand testCMD = new SqlCommand("POCountForUser", connection);
                        testCMD.CommandType = CommandType.StoredProcedure;

                        testCMD.Parameters.Add(new SqlParameter("@POId", System.Data.SqlDbType.BigInt, 50) { Value = POID });
                        testCMD.Parameters.Add(new SqlParameter("@UserId", System.Data.SqlDbType.BigInt, 50) { Value = (_login.EmployeeID).ToString() });

                        SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(testCMD);

                        DataSet dataSet = new DataSet();
                        sqlDataAdapter.Fill(dataSet);

                        int counter = 0;
                        while (counter < dataSet.Tables[0].Rows.Count)
                        {
                            isFound = Convert.ToInt32(dataSet.Tables[0].Rows[counter]["Count"]);
                            counter++;
                        }
                        dataSet.Dispose();
                    }
                    if (isFound != null)
                    {
                        QuoteComparer qC = new QuoteComparer(_login, POID);
                        qC.Show();
                        this.Close();
                    }
                    else
                    {
                        MessageBox.Show("You are not authorised to view the PO " + POID,
                                       "Order Management System",
                                           MessageBoxButton.OK,
                                               MessageBoxImage.Error);
                    }
                }
                else
                {
                    ViewPurchaseOrders vPO = new ViewPurchaseOrders(_login);
                    vPO.Show();
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured during PO search " + ex.Message,
                                   "Order Management System",
                                       MessageBoxButton.OK,
                                           MessageBoxImage.Error);
            }
        }

        private void btn_search_indent_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //log.Info("Search Indent activated...");
                if (txt_indent_no.Text != null && txt_indent_no.Text != "")
                {
                    long indentNo = long.Parse(txt_indent_no.Text);

                    //var idFound = (from i in orderManagementContext.IndentApproval
                    //               where i.IndentId == indentNo
                    //               select i).FirstOrDefault();

                    int? isFound = null;

                    using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlConnection"].ToString()))
                    {
                        connection.Open();

                        SqlCommand testCMD = new SqlCommand("IndentCountForUser", connection);
                        testCMD.CommandType = CommandType.StoredProcedure;

                        testCMD.Parameters.Add(new SqlParameter("@IndentID", System.Data.SqlDbType.BigInt, 50) { Value = indentNo });
                        testCMD.Parameters.Add(new SqlParameter("@UserId", System.Data.SqlDbType.BigInt, 50) { Value = (_login.EmployeeID).ToString() });

                        SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(testCMD);

                        DataSet dataSet = new DataSet();
                        sqlDataAdapter.Fill(dataSet);

                        int counter = 0;
                        while (counter < dataSet.Tables[0].Rows.Count)
                        {
                            isFound = Convert.ToInt32(dataSet.Tables[0].Rows[counter]["Count"]);
                            counter++;
                        }
                        dataSet.Dispose();
                    }

                    if (isFound != null)
                    {
                        Indent indent = new Indent(_login, indentNo);
                        indent.Show();
                        this.Close();
                    }
                    else
                    {
                        MessageBox.Show("You are not authorised to view the indent " + indentNo,
                                  "Order Management System",
                                      MessageBoxButton.OK,
                                          MessageBoxImage.Error);
                    }
                }
                else
                {
                    //log.Info("Opening View Indent screen...");
                    ViewIndents viewIndent = new ViewIndents(_login);
                    viewIndent.Show();
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occured while fetching user information. " + ex.Message,
                                   "Order Management System",
                                       MessageBoxButton.OK,
                                           MessageBoxImage.Error);
                //log.Error("View Indent Click error : " + ex.StackTrace);
            }
        }
    }
}