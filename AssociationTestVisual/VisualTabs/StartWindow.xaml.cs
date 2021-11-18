using System.Windows;
using ExcelHelper;
using System.Collections;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace AssociationTestVisual.VisualTabs
{
    /// <summary>
    /// Логика взаимодействия для Window2.xaml
    /// </summary>


    public partial class StartWindow : Window
    {
        public GroupsList Groups = new GroupsList();
        public StartWindow()
        {
            InitializeComponent();
            //Groups.List.Add("TestGroup");
            //Groups.Save();
            Groups.Load();
            foreach (var v in Groups.List) { GROUPBox.Items.Add(v); }
            GROUPBox.SelectedIndex = 0;
            
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            await Task.Factory.StartNew(() =>
            {
            GLOBALS.Eww = new ExcelWorker();
            Dispatcher.Invoke(() => { ContinueButton.IsEnabled = true; });
            });
        }

        private void ContinueButton_Click(object sender, RoutedEventArgs e)
        {
            //if (FIOBox.Text.Length == 0) { MessageBox.Show("Поле ФИО или ID не должно быть пустым!!!"); return; }
           // else
            //if (GROUPBox.Text.Length == 0) { MessageBox.Show("Вы должны выбрать группу тестируемого!!!"); return; }
            GLOBALS.GetPerson = new PersonResult() { Name = FIOBox.Text, Group = GROUPBox.Text };
            Input inputWindow = new Input();
            inputWindow.Show();
            this.Close();
        }


        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            //Dispatcher.Invoke(()=> { GLOBALS.Eww?.Close(); });
        }
    }
}
