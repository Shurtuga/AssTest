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
    public partial class CatsSortWindow : Window
    {
        PersonResult res;
        ExcelWorker Eww;
        List<ListBox> lists;
        List<ContextMenu> menus = new List<ContextMenu>();
        List<StackPanel> cats;
        List<WordInfo> wordInfos = new List<WordInfo>();
        List<TabItem> tabs;
        WordsList words;
        public CatsSortWindow(PersonResult r, List<WordInfo> w, ExcelWorker ew, WordsList wo)
        {
            res = r;
            wordInfos = w;
            Eww = ew;
            words = wo;
            InitializeComponent();
            tabs = new List<TabItem> { TIFir, TISecond, TIThird, TIFourth, TIFifth, TISixth };
            cats = new List<StackPanel> { FirCategory, SecondCategory, ThirdCategory, FourthCategory, FifthCategory, SixthCategory };
            lists = new List<ListBox> { FirUnsortedWordsList, SecondUnsortedWordsList, ThirdUnsortedWordsList, FourthUnsortedWordsList, FifthUnsortedWordsList, SixthUnsortedWordsList };

            for (int i = 0; i < 6; i++)
            {
                ContextMenu menu = new ContextMenu();
                tabs[i].Header = words.semanticMeanings[i].Word;
                for (int j = 0; j < WordsList.assTypes.Count; j++)
                {
                    MenuItem mi = new MenuItem() { Name = "ItName_" + i.ToString() + "_" + j.ToString(), Header = WordsList.assTypes[j] };
                    mi.Click += Mi_Click;
                    menu.Items.Add(mi);
                    StackPanel sp = new StackPanel() { Margin = new Thickness(10, 2, 10, 2), VerticalAlignment = VerticalAlignment.Stretch, Height = Double.NaN };
                    sp.Children.Add(new Button() { Content = WordsList.assTypes[j], FontSize = 20 });
                    cats[i].Children.Add(sp);
                }
                menus.Add(menu);
            }

            void Mi_Click(object sender, RoutedEventArgs e)
            {
                MenuItem mnu = sender as MenuItem;
                string[] s = ((MenuItem)sender).Name.Split('_');
                //добавление в выбранную категорию!!!
                ((StackPanel)cats[int.Parse(s[1])].Children[int.Parse(s[2])]).Children.Add(new TextBlock() { Text = ((TextBlock)((ContextMenu)mnu.Parent).PlacementTarget).Text, FontSize = 15, VerticalAlignment = VerticalAlignment.Center, HorizontalAlignment = HorizontalAlignment.Center });
                lists[int.Parse(s[1])].Items.Remove((TextBlock)((ContextMenu)mnu.Parent).PlacementTarget);
            }

            foreach (var v in wordInfos)
            {
                int i = words.semanticMeanings.FindIndex(t => t.Word.ToLower() == v.Association.ToLower());
                if (v.FAss!=-1) { ((StackPanel)cats[i].Children[v.FAss-1]).Children.Add(new TextBlock() { Text = v.Word }); }
                else lists[i].Items.Add(new TextBlock() { Text = v.Word, ContextMenu = menus[i] });
            }

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i<6; i++)
            {
                if (lists[i].Items.Count !=0) { MessageBox.Show("Необходимо распределить ассоциации к слову "+tabs[i].Header.ToString()); return; }
            }
            ResultWindow rw = new ResultWindow(res);//сюда передать итоговый резалт
            rw.Show();
            this.Close();
        }
    }
}
