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
using System.IO;

namespace AssociationTestVisual.VisualTabs
{
    /// <summary>
    /// Логика взаимодействия для Window1.xaml
    /// </summary>
    public partial class Input : Window
    {
        PersonResult results;
        WordsList Words = new WordsList();
        List<TextBox> inpts;
        List<ListBox> lists;
        List<ContextMenu> menus = new List<ContextMenu>();
        List<StackPanel> cats;
        List<WordInfo> wordInfos = new List<WordInfo>();
        List<TabItem> tabs;
        //ExcelWorker Eww = new ExcelWorker(System.IO.Path.GetFullPath(@"..\..\..\ExcelHelper\ExcelTables\Частоты.xlsx"));
        ExcelWorker Eww;

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
            results = res;
            tabs = new List<TabItem>{ TIFir, TISecond, TIThird, TIFourth, TIFifth, TISixth };
            cats = new List<StackPanel> { FirCategory, SecondCategory, ThirdCategory, FourthCategory, FifthCategory, SixthCategory };
            inpts = new List<TextBox> { FirWordsInput, SecondWordsInput, ThirdWordsInput, FourthWordsInput, FifthWordsInput, SixthWordsInput };
            lists = new List<ListBox> { FirUnsortedWordsList, SecondUnsortedWordsList, ThirdUnsortedWordsList, FourthUnsortedWordsList, FifthUnsortedWordsList, SixthUnsortedWordsList };
            for (int i = 0; i < 6; i++)
            {
                ContextMenu menu = new ContextMenu();
                tabs[i].Header = Words.semanticMeanings[i].Word;
                for (int j = 0; j < Words.semanticMeanings[i].Meanings.Count; j++)
                {
                    MenuItem mi = new MenuItem() { Name = "ItName_" + i.ToString() + "_" + j.ToString(), Header = Words.semanticMeanings[i].Meanings[j] };
                    mi.Click += Mi_Click;
                    menu.Items.Add(mi);
                    StackPanel sp = new StackPanel() { Margin = new Thickness(10, 2, 10, 2), VerticalAlignment = VerticalAlignment.Stretch, Height = Double.NaN };
                    sp.Children.Add(new Button() { Content = Words.semanticMeanings[i].Meanings[j], FontSize = 20 });
                    cats[i].Children.Add(sp);
                }
                menus.Add(menu);
            }
        }

        private void Mi_Click(object sender, RoutedEventArgs e)
        {
            MenuItem mnu = sender as MenuItem;
            string[] s = ((MenuItem)sender).Name.Split('_');
            //добавление в выбранную категорию
            ((StackPanel)cats[int.Parse(s[1])].Children[int.Parse(s[2])]).Children.Add(new TextBlock() { Text = ((TextBlock)((ContextMenu)mnu.Parent).PlacementTarget).Text, FontSize = 15, VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Center });
            lists[int.Parse(s[1])].Items.Remove((TextBlock)((ContextMenu)mnu.Parent).PlacementTarget);
        }

        private async void EnterWord(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Return)
            {
                //если незнакомое слово
                int i = inpts.IndexOf((TextBox)sender);

                string word = inpts[i].Text;
                inpts[i].Text = "";

                //var gr = Eww.GetWordAsync(inpts[i].Text, tabs[i].Header.ToString());
                //var r = await gr;

                wordInfos.Add(await Eww.GetWordAsync(word, tabs[i].Header.ToString()));

                //wordInfos.Add(r);
                //wordInfos.Add(Eww.GetWord(inpts[i].Text, tabs[i].Header.ToString()));

                if (wordInfos.Last().FSem !=-1)
                {
                    //добавление в выбранную категорию
                    ((StackPanel)cats[i].Children[wordInfos.Last().FSem-1]).Children.Add(new TextBlock() { Text = wordInfos.Last().Word });
                }
                else lists[i].Items.Add(new TextBlock() { Text = word, ContextMenu = menus[i] });
                //inpts[i].Text = "";
                //иначе получаем индекс категории, добавляем в i-тый стекбокс
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i<6; i++) 
            {
                if (lists[i].Items.Count !=0) { MessageBox.Show("Необходимо распределить ассоциации к слову "+tabs[i].Header.ToString()); return; } 
            }
            CatsSortWindow CatsSort = new CatsSortWindow(results, wordInfos, Eww, Words);
            CatsSort.Show();
            //Eww.Close();
            this.Close();
        }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            await Task.Factory.StartNew(() =>
            {
                Eww = new ExcelWorker();
            });
        }
    }
}
