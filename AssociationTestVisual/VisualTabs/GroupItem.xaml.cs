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
    /// Логика взаимодействия для GroupItem.xaml
    /// </summary>
    public partial class GroupItem : ComboBoxItem
    {
        public GroupItem(string name)
        {
            InitializeComponent();
            GroupName.Text = name;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            GLOBALS.Groups.List.Remove(GroupName.Text);
            ((ComboBox)Parent).Items.Remove(this);
        }

        public override string ToString()
        {
            return GroupName.Text;
        }
    }
}
