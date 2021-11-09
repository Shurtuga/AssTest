using System.Windows;
using ExcelHelper;
using System.Collections;
using System.Collections.Generic;

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

        private void ContinueButton_Click(object sender, RoutedEventArgs e)
        {
            if (FIOBox.Text.Length == 0) { MessageBox.Show("Поле ФИО или ID не должно быть пустым!!!"); return; }
           // else
            //if (GROUPBox.Text.Length == 0) { MessageBox.Show("Вы должны выбрать группу тестируемого!!!"); return; }
            PersonResults pr = new PersonResults() { Name = FIOBox.Text, Group = GROUPBox.Text };
            Input inputWindow = new Input(pr);
            inputWindow.Show();
            this.Close();
        }
    }
}
