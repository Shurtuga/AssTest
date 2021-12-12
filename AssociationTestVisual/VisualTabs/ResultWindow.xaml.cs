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

        private void SaveRef(object sender, RoutedEventArgs e)
        {
            foreach (var v in GLOBALS.WordInfos) { GLOBALS.Eww.AddWord(v); }

            GLOBALS.Eww.ResultReferencePhase();

            GLOBALS.Eww.SaveResultRef(GLOBALS.GetPerson, GLOBALS.WordInfos);
            SaveButton.IsEnabled = false;
        }

        private void SaveAllButton_Click(object sender, RoutedEventArgs e)
        {
            GLOBALS.Eww.ResultPhase();

            GLOBALS.Eww.SaveAllResults();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            
        }

        private void RestartButton_Click(object sender, RoutedEventArgs e)
        {
            if (SaveButton.IsEnabled)
            {
                if (MessageBox.Show("Вы не сохранили результат! Сохранить?", "Внимание!", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    GLOBALS.Eww.SaveResult(GLOBALS.GetPerson);
                }
            }
            StartWindow sw = new StartWindow();
            sw.Show();
            this.Close();
        }

       
    }
}
