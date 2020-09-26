using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Animation;

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
