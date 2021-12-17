using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
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

            var personres = GLOBALS.Eww.Calculate(GLOBALS.GetPerson.Name, GLOBALS.GetPerson.Group, GLOBALS.WordInfos);
            FIOBox.Text = personres.Name;
            GroupBox.Text = personres.Group;
            SpeedBox.Text = personres.Speed.ToString();
            FassBox.Text = personres.FAss.ToString();
            FsemBox.Text = personres.FSem.ToString();
        }

        private void SaveRef(object sender, RoutedEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                foreach (var v in GLOBALS.WordInfos) { GLOBALS.Eww.AddWord(v); }

                GLOBALS.Eww.ResultReferencePhase();

                GLOBALS.Eww.SaveResultRef(GLOBALS.GetPerson, GLOBALS.WordInfos);
                SaveButton.IsEnabled = false;
            });
        }

        private async void SaveAllButton_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Вы планируете завершить тестирование.\nЕсли вы нажмёте \"Да\", то вы потеряете возможность продолжать проверку данного тестирования.", "Внимание!", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
            {
                SaveAllButton.IsEnabled = false;
                Gbox.Visibility = Visibility.Visible;
                await Task.Factory.StartNew(() =>
                {
                    GLOBALS.Eww.ResultPhase();
                    GLOBALS.Eww.SaveAllResults();
                    Dispatcher.Invoke(() => Gbox.Visibility = Visibility.Hidden);
                });
                GLOBALS.Groups.List.Clear();
            }


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
                    //GLOBALS.Eww.SaveResultRef(GLOBALS.GetPerson, GLOBALS.WordInfos);
                    SaveRef(sender, e);
                }
            }
            StartWindow sw = new StartWindow();
            sw.Show();
            this.Close();
        }


    }
}
