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
        public Input()
        {
            InitializeComponent();
            //Words.semanticMeanings.Add(new AssociationWord() { Word = "Клетка", Meanings = new List<string>() {"Заграждение", "Часть организма", "Геометрическая фигура", "Лестничная клетка", "Образное сравнение", "Имена собственные", "Рифмовки" } });
            //Words.semanticMeanings.Add(new AssociationWord() { Word = "Порог", Meanings = new List<string>() { "Брус под дверью", "Начало, рубеж, граница", "Вода", "Рифмовка", "Часть музыкального инструмента", "Препятствие", "Сложнопонимаемая ассоциация" } });
            //Words.semanticMeanings.Add(new AssociationWord() { Word = "Лист", Meanings = new List<string>() { "Часть растения", "Кусок материала", "Единица печатного объёма", "Имя собственное", "Рифмовка"} });
            //Words.semanticMeanings.Add(new AssociationWord() { Word = "Ключ", Meanings = new List<string>() { "Родник, источник", "Отпирающее замки", "Отвинчивающее гайки", "Ключ к отгадке", "Шифр, код", "Имя собственное", "Музыкальная грамота" } });
            //Words.semanticMeanings.Add(new AssociationWord() { Word = "Язык", Meanings = new List<string>() { "Орган вкуса", "Часть колокола", "Деталь обуви", "Словесное средство", "Невербалика", "Народ, нация", "Формальный язык", "Выпечка, пирожное","Блюдо, продукт" } });
            //Words.semanticMeanings.Add(new AssociationWord() { Word = "Дробь", Meanings = new List<string>() { "Мелкие ядрышки", "Звуки", "Математика", "Обозначение номера дома" } });
            //Words.Save();
            GLOBALS.Words = new WordsList();
            GLOBALS.WordInfos = new List<WordInfo>();
            GLOBALS.Words.Load();
            int i = 0;
            foreach (var v in GLOBALS.Words.semanticMeanings) { Tabs.Items.Add(new MyTabItem(v.Word, i).Tab);i++; }
            
        }


        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i<Tabs.Items.Count; i++)
            {
                if (((MyTabItem)Tabs.Items[i]).UnsortedList.Items.Count !=0||((MyTabItem)Tabs.Items[i]).UnsortedAssList.Items.Count !=0) { MessageBox.Show("Необходимо распределить ассоциации к слову "+((MyTabItem)Tabs.Items[i]).InWord); return; }
            }

            ResultWindow rw = new ResultWindow();//сюда передать итоговый резалт
            rw.Show();
            this.Close();
        }

        
    }
}
