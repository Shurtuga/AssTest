using System;
using System.Collections.Generic;
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
            //    GLOBALS.Eww?.Close();
        }
    }
    public static class GLOBALS
    {
        public static ExcelHelper.ExcelWorker Eww { get; set; }
        public static ExcelHelper.PersonResult GetPerson { get; set; }

        public static VisualTabs.WordsList Words { get; set; }
        public static List<ExcelHelper.WordInfo> WordInfos { get; set; }
    }
}
