using System;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace AssociationTestVisual
{
    /// <summary>
    /// Логика взаимодействия для App.xaml
    /// </summary>
    public partial class App : Application
    {
        private void Application_Exit(object sender, ExitEventArgs e)
        {
            GLOBALS.Eww?.Close();

            
        }

        private void Application_Startup(object sender, StartupEventArgs e)
        {
            //GLOBALS.Eww = new ExcelHelper.ExcelWorker();
            App.Current.DispatcherUnhandledException += (s, ed) =>
            {
                if (MessageBox.Show($"Исключение: {ed.Exception.Message}\n{ed.Exception.Source}\n{ed.Exception.StackTrace}\n{ed.Exception.HelpLink}") == MessageBoxResult.OK)
                {
                }
            };
        }
    }
}
