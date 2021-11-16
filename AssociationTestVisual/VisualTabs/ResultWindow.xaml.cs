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
    /// Логика взаимодействия для ResultWindow.xaml
    /// </summary>
    /// 
    public partial class ResultWindow : Window
    {
        PersonResult result;
        public ResultWindow(PersonResult res)
        {
            InitializeComponent();
            result = res;
            FIOBox.Text = result.Name;
            GroupBox.Text = result.Group;
            SpeedBox.Text = result.Speed.ToString();
            OriginalityBox.Text = result.Originality.ToString();
            FassBox.Text = result.FAss.ToString();
            FsemBox.Text = result.FSem.ToString();
        }
    }
}
