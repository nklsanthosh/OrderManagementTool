using log4net;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace OrderManagementTool
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        ILog log = LogManager.GetLogger(typeof(App));

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            log.Info("Initialising ...");
            //initialize the splash screen and set it as the application main window
            var splashScreen = new SplashScreen();
            this.MainWindow = splashScreen;
            splashScreen.Show();

            //in order to ensure the UI stays responsive, we need to
            //do the work on a different thread
            Task.Factory.StartNew(() =>
            {
                //simulate some work being done
                System.Threading.Thread.Sleep(3000);

                //since we're not on the UI thread
                //once we're done we need to use the Dispatcher
                //to create and show the main window
                this.Dispatcher.Invoke(() =>
                {
                    //initialize the main window, set it as the application main window
                    //and close the splash screen
                    var loginWindow = new MainWindow();
                    this.MainWindow = loginWindow;
                    loginWindow.Show();
                    splashScreen.Close();

                });
            });
            
        }
    }
}
