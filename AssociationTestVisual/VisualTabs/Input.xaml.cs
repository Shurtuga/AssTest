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
    /// Логика взаимодействия для Window1.xaml
    /// </summary>
    public partial class Input : Window
    {
        public PersonResults results;
        WordsList Words = new WordsList();
        public Input(PersonResults res)
        {
            InitializeComponent();
            //Words.semanticMeanings.Add(new AssociationWord() { Word = "Клетка", Meanings = new List<string>() {"Заграждение", "Часть организма", "Геометрическая фигура", "Лестничная клетка", "Образное сравнение", "Имена собственные", "Рифмовки" } });
            //Words.semanticMeanings.Add(new AssociationWord() { Word = "Порог", Meanings = new List<string>() { "Брус под дверью", "Начало, рубеж, граница", "Вода", "Рифмовка", "Часть музыкального инструмента", "Препятствие", "Сложнопонимаемая ассоциация" } });
            //Words.semanticMeanings.Add(new AssociationWord() { Word = "Лист", Meanings = new List<string>() { "Часть растения", "Кусок материала", "Единица печатного объёма", "Имя собственное", "Рифмовка"} });
            //Words.semanticMeanings.Add(new AssociationWord() { Word = "Ключ", Meanings = new List<string>() { "Родник, источник", "Отпирающее замки", "Отвинчивающее гайки", "Ключ к отгадке", "Шифр, код", "Имя собственное", "Музыкальная грамота" } });
            //Words.semanticMeanings.Add(new AssociationWord() { Word = "Язык", Meanings = new List<string>() { "Орган вкуса", "Часть колокола", "Деталь обуви", "Словесное средство", "Невербалика", "Народ, нация", "Формальный язык", "Выпечка, пирожное","Блюдо, продукт" } });
            //Words.semanticMeanings.Add(new AssociationWord() { Word = "Дробь", Meanings = new List<string>() { "Мелкие ядрышки", "Звуки", "Математика", "Обозначение номера дома" } });
            //Words.Save();
            Words.Load();
            Tabs.Items.Clear();
            foreach (var v in Words.semanticMeanings)
            {
                TI = (TabItem) FindResource("TICOPY");
                TI.Header = v.Word;
                foreach (var m in v.Meanings) 
                {
                    StackPanel sp = new StackPanel() { Margin = new Thickness(10, 10, 10, 10), VerticalAlignment = VerticalAlignment.Stretch };
                    sp.Children.Add(new TextBlock() { Text = m });
                    ((StackPanel)TI.FindName("CCategory")).Children.Add(sp);
                }

                Tabs.Items.Add(TI);
            }
            results = res;
        }
        public void WordEntered()
        {
            //    UnsortedWordsList.Items.Add(new TextBlock() { Text = WordsInput.Text});
            //    WordsInput.Clear();
        }
        private void EnterWord(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Return)
            {
                WordEntered();
            }
        }
    }
}
