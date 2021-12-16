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
        ContextMenu SemMenu;
        ContextMenu AssMenu;
        public int Numer = -1;
        public MyTabItem(string word, int n)
        {
            InitializeComponent();
            Numer = n;
            Tab.Header = word;
            InWord = word;
            SemMenu = new ContextMenu();
            AssMenu = new ContextMenu();
            SemInit();
            AssInit();
        }
        #region Semantics

        private void SemInit()
        {
            for (int j = 0; j < GLOBALS.Words.semanticMeanings[Numer].Meanings.Count; j++)
            {
                MenuItem mi = new MenuItem() { Name = "ItName_" + Numer.ToString() + "_" + j.ToString(), Header = GLOBALS.Words.semanticMeanings[Numer].Meanings[j] };
                mi.Click += SemMi_Click;
                SemMenu.Items.Add(mi);
                StackPanel sp = new StackPanel() { Margin = new Thickness(10, 2, 10, 2), VerticalAlignment = VerticalAlignment.Stretch, Height = Double.NaN };
                sp.Children.Add(new Button() { Content = GLOBALS.Words.semanticMeanings[Numer].Meanings[j], FontSize = 20 });
                Category.Children.Add(sp);
            }
            MenuItem mu = new MenuItem() { Name = "Del", Header = "Удалить" };
            mu.Click += SemMi_Click;
            SemMenu.Items.Add(mu);
        }

        private void SemMi_Click(object sender, RoutedEventArgs e)
        {
            MenuItem mnu = sender as MenuItem;
            string[] s = ((MenuItem)sender).Name.Split('_');
            if (((MenuItem)sender).Name=="Del")
            {
                GLOBALS.WordInfos.RemoveAt(GLOBALS.WordInfos.FindIndex(t => t.Word == ((TextBlock)((ContextMenu)mnu.Parent).PlacementTarget).Text));
                UnsortedList.Items.Remove((TextBlock)((ContextMenu)mnu.Parent).PlacementTarget);
                for (int i = 0; i<UnsortedAssList.Items.Count; i++)
                {
                    if (((TextBlock)(UnsortedAssList.Items[i])).Text == ((TextBlock)((ContextMenu)mnu.Parent).PlacementTarget).Text)
                    {
                        UnsortedAssList.Items.RemoveAt(i);
                        break;
                    }
                }
                return;
            }
            //добавление в выбранную категорию
            ((StackPanel)Category.Children[int.Parse(s[2])]).Children.Add(new TextBlock() { Text = ((TextBlock)((ContextMenu)mnu.Parent).PlacementTarget).Text, FontSize = 20, VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Center });
            UnsortedList.Items.Remove((TextBlock)((ContextMenu)mnu.Parent).PlacementTarget);
            for (int i = 0; i< GLOBALS.WordInfos.Count; i++)
            {
                if (GLOBALS.WordInfos[i].Word == ((TextBlock)((ContextMenu)mnu.Parent).PlacementTarget).Text)
                {
                    GLOBALS.WordInfos[i].FSem = int.Parse(s[2])+1;
                    break;
                }
            }
        }

        private async void SemEnterWord(object sender, KeyEventArgs e)
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
                    ((StackPanel)AssCategory.Children[GLOBALS.WordInfos.Last().FSem-1]).Children.Add(new TextBlock() { Text =GLOBALS.WordInfos.Last().Word });
                }
                else
                {
                    UnsortedList.Items.Add(new TextBlock() { Text = word, ContextMenu = SemMenu });
                    UnsortedAssList.Items.Add(new TextBlock() { Text = word, ContextMenu = AssMenu });
                };
                //иначе получаем индекс категории, добавляем в i-тый стекбокс
            }
        }
        #endregion

        #region Association

        public void AssInit()
        {
            for (int j = 0; j < WordsList.assTypes.Count; j++)
            {
                MenuItem mi = new MenuItem() { Name = "ItName_" + Numer.ToString() + "_" + j.ToString(), Header = WordsList.assTypes[j] };
                mi.Click += AssMi_Click;
                AssMenu.Items.Add(mi);
                StackPanel sp = new StackPanel() { Margin = new Thickness(10, 2, 10, 2), VerticalAlignment = VerticalAlignment.Stretch, Height = Double.NaN };
                sp.Children.Add(new Button() { Content = WordsList.assTypes[j], FontSize = 20 });
                AssCategory.Children.Add(sp);
            }
            MenuItem mu = new MenuItem() { Name = "Del", Header = "Удалить" };
            mu.Click += AssMi_Click;
            AssMenu.Items.Add(mu);
        }

        void AssMi_Click(object sender, RoutedEventArgs e)
        {
            MenuItem mnu = sender as MenuItem;
            if (((MenuItem)sender).Name=="Del")
            {
                GLOBALS.WordInfos.RemoveAt(GLOBALS.WordInfos.FindIndex(t => t.Word == ((TextBlock)((ContextMenu)mnu.Parent).PlacementTarget).Text));
                UnsortedAssList.Items.Remove((TextBlock)((ContextMenu)mnu.Parent).PlacementTarget);
                for (int i = 0; i<UnsortedList.Items.Count; i++)
                {
                    if (((TextBlock)(UnsortedList.Items[i])).Text == ((TextBlock)((ContextMenu)mnu.Parent).PlacementTarget).Text)
                    {
                        UnsortedList.Items.RemoveAt(i);
                        break;
                    }
                }
                return;
            }
            string[] s = ((MenuItem)sender).Name.Split('_');
            //добавление в выбранную категорию!!!
            ((StackPanel)AssCategory.Children[int.Parse(s[2])]).Children.Add(new TextBlock() { Text = ((TextBlock)((ContextMenu)mnu.Parent).PlacementTarget).Text, FontSize = 20, VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Center });
            UnsortedAssList.Items.Remove((TextBlock)((ContextMenu)mnu.Parent).PlacementTarget);

            for (int i = 0; i< GLOBALS.WordInfos.Count; i++)
            {
                if (GLOBALS.WordInfos[i].Word == ((TextBlock)((ContextMenu)mnu.Parent).PlacementTarget).Text)
                {
                    GLOBALS.WordInfos[i].FAss = int.Parse(s[2])+1;
                    break;
                }
            }
        }
        #endregion
    }
}
