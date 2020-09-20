using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace OrderManagementTool
{
    /// <summary>
    /// Interaction logic for RadialProgressBar.xaml
    /// </summary>
    public partial class RadialProgressBar : UserControl
    {
        static RadialProgressBar()
        {
            Timeline.DesiredFrameRateProperty.OverrideMetadata(
                 typeof(Timeline),
                    new FrameworkPropertyMetadata { DefaultValue = 20 });
        }
        public RadialProgressBar()
        {
            InitializeComponent();

           
        }
    }
}
