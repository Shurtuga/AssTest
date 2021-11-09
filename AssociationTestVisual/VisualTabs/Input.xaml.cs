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
        public PersonResult results;
        WordsList Words = new WordsList();
        public Input(PersonResult res)
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
            var tabs = new[] { TIFir, TISecond, TIThird, TIFourth, TIFifth, TISixth };
            var cats = new[] { FirCategory, SecondCategory, ThirdCategory, FourthCategory, FifthCategory, SixthCategory };
            for (int i = 0; i < 6; i++)
            {
                tabs[i].Header = Words.semanticMeanings[i].Word;
                for (int j = 0; j < Words.semanticMeanings[i].Meanings.Count;j++)
                {
                    StackPanel sp = new StackPanel() {Margin = new Thickness(10,2,10,2), VerticalAlignment = VerticalAlignment.Stretch, Height = Double.NaN };
                    sp.Children.Add(new Button() { Content = Words.semanticMeanings[i].Meanings[j], FontSize = 20 });
                    cats[i].Children.Add(sp);
                }
            }
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
