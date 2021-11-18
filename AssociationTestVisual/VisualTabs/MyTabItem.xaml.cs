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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace AssociationTestVisual.VisualTabs
{
    /// <summary>
    /// Логика взаимодействия для MyTabItem.xaml
    /// </summary>
    public partial class MyTabItem : TabItem
    {
        public string InWord;
        ContextMenu menu;
        public int Numer = -1;
        public MyTabItem(string word, int n)
        {
            InitializeComponent();
            Numer = n;
            Tab.Header = word;
            InWord = word;
            menu = new ContextMenu();
            for (int j = 0; j < GLOBALS.Words.semanticMeanings[Numer].Meanings.Count; j++)
            {
                MenuItem mi = new MenuItem() { Name = "ItName_" + Numer.ToString() + "_" + j.ToString(), Header = GLOBALS.Words.semanticMeanings[Numer].Meanings[j] };
                mi.Click += Mi_Click;
                menu.Items.Add(mi);
                StackPanel sp = new StackPanel() { Margin = new Thickness(10, 2, 10, 2), VerticalAlignment = VerticalAlignment.Stretch, Height = Double.NaN };
                sp.Children.Add(new Button() { Content = GLOBALS.Words.semanticMeanings[Numer].Meanings[j], FontSize = 20 });
                Category.Children.Add(sp);
            }
        }
        private void Mi_Click(object sender, RoutedEventArgs e)
        {
            MenuItem mnu = sender as MenuItem;
            string[] s = ((MenuItem)sender).Name.Split('_');
            //добавление в выбранную категорию
            ((StackPanel)Category.Children[int.Parse(s[2])]).Children.Add(new TextBlock() { Text = ((TextBlock)((ContextMenu)mnu.Parent).PlacementTarget).Text, FontSize = 15, VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Center });
            UnsortedList.Items.Remove((TextBlock)((ContextMenu)mnu.Parent).PlacementTarget);
        }

        private async void EnterWord(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Return)
            {
                //если незнакомое слово

                string word = WordsInput.Text;
                WordsInput.Text = "";
                //var gr = Eww.GetWordAsync(inpts[i].Text, tabs[i].Header.ToString());
                //var r = await gr;
                GLOBALS.WordInfos.Add(await GLOBALS.Eww.GetWordAsync(word, InWord));
                
                //wordInfos.Add(r);
                //wordInfos.Add(Eww.GetWord(inpts[i].Text, tabs[i].Header.ToString()));

                if (GLOBALS.WordInfos.Last().FSem !=-1)
                {
                    //добавление в выбранную категорию
                    ((StackPanel)Category.Children[GLOBALS.WordInfos.Last().FSem-1]).Children.Add(new TextBlock() { Text = GLOBALS.WordInfos.Last().Word });
                }
                else UnsortedList.Items.Add(new TextBlock() { Text = word, ContextMenu = menu });
                //иначе получаем индекс категории, добавляем в i-тый стекбокс
            }
        }
    }
}
