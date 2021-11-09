using System.Windows;
using ExcelHelper;

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
        }

        private void ContinueButton_Click(object sender, RoutedEventArgs e)
        {
            if (FIOBox.Text.Length == 0) { MessageBox.Show("Поле ФИО или ID не должно быть пустым!!!"); return; }
           // else
            //if (GROUPBox.Text.Length == 0) { MessageBox.Show("Вы должны выбрать группу тестируемого!!!"); return; }
            PersonResult pr = new PersonResult() { Name = FIOBox.Text, Group = GROUPBox.Text };
            Input inputWindow = new Input(pr);
            inputWindow.Show();
            this.Close();
        }
    }
}
