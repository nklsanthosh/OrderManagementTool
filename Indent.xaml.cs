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

namespace OrderManagementTool
{
    /// <summary>
    /// Interaction logic for Indent.xaml
    /// </summary>
    public partial class Indent : Window
    {

        private readonly List<string> itemName;
        public Indent()
        {
            InitializeComponent();
            OrderManagementContext orderManagementContext = new OrderManagementContext();
            itemName = (from i in orderManagementContext.ItemCategory
                        select i.ItemCategoryName).Distinct().ToList();
            grid_indentdata.DataContext = itemName;
        }

        private void DataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }
}
