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
        public Input( PersonResults res)
        {
            InitializeComponent();
            results = res;
        }
        public void WordEntered()
        {
            UnsortedWordsList.Items.Add(new TextBlock() { Text = WordsInput.Text});
            WordsInput.Clear();
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
