using System.Windows;
using ExcelHelper;
using System.Collections;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace AssociationTestVisual.VisualTabs
{
    /// <summary>
    /// Логика взаимодействия для Window2.xaml
    /// </summary>


    public partial class StartWindow : Window
    {
        public StartWindow()
       {
            InitializeComponent();
            GLOBALS.Groups = new GroupsList();
            GLOBALS.Groups.Load();
            if (GLOBALS.Eww==null)
            {
                GLOBALS.Eww = new ExcelWorker();
                GLOBALS.Eww.ExcelLoaded+=(s, e) => { Dispatcher.Invoke(() => { ContinueButton.IsEnabled = true; ContinueButton.Content = "Продолжить"; }); };
            }
            else { ContinueButton.IsEnabled = true; ContinueButton.Content = "Продолжить"; }
            GLOBALS.Eww.InputPhaseAsync();
            FillGroups();
        }


        public void FillGroups() 
        {
            foreach (var group in GLOBALS.Groups.List) 
            {
                GROUPBox.Items.Add(new GroupItem(group));
            }
            Button btn = new Button() {Content = "Добавить группу", HorizontalAlignment = HorizontalAlignment.Stretch, VerticalAlignment = VerticalAlignment.Stretch, Height = 35 };
            btn.Click+=AddItem;
            GROUPBox.Items.Add(btn);

        }

        private void AddItem(object sender, RoutedEventArgs e)
        {
            string s = Microsoft.VisualBasic.Interaction.InputBox("Введите название новой группы", "Добавление группы");
            if (s!="")
            {
                GroupItem gi = new GroupItem(s);
                GROUPBox.Items.Insert(GROUPBox.Items.Count-1, gi);
                GLOBALS.Groups.List.Add(s);
                GROUPBox.IsDropDownOpen = true;
            }
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
        }

        private void ContinueButton_Click(object sender, RoutedEventArgs e)
        {
            if (FIOBox.Text.Length == 0) { MessageBox.Show("Поле ФИО или ID не должно быть пустым!!!"); return; }
            else
            if (GROUPBox.SelectedIndex<0) { MessageBox.Show("Вы должны выбрать группу тестируемого!!!"); return; }
            GLOBALS.GetPerson = new PersonResult() { Name = FIOBox.Text, Group = GROUPBox.SelectedItem.ToString() };
            Input inputWindow = new Input();
            inputWindow.Show();
            this.Close();
        }


        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
        }

        
    }
}
