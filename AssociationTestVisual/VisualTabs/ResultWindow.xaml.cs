using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using ExcelHelper;

namespace AssociationTestVisual.VisualTabs
{
    /// <summary>
    /// Логика взаимодействия для ResultWindow.xaml
    /// </summary>
    /// 
    public partial class ResultWindow : Window
    {
        public ResultWindow()
        {
            InitializeComponent();
            FIOBox.Text = GLOBALS.GetPerson.Name;
            GroupBox.Text = GLOBALS.GetPerson.Group;
			SpeedBox.Text = GLOBALS.GetPerson.Speed.ToString();
            OriginalityBox.Text = GLOBALS.GetPerson.Originality.ToString();
            FassBox.Text = GLOBALS.GetPerson.FAss.ToString();
            FsemBox.Text = GLOBALS.GetPerson.FSem.ToString();
            
            
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            GLOBALS.Eww.SaveResult(GLOBALS.GetPerson);
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            foreach (var v in GLOBALS.WordInfos) { await GLOBALS.Eww.AddWordAsync(v); }
        }
    }
}
