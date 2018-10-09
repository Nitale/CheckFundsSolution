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

namespace CheckFunds
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            CheckFunds.Reader reader = new CheckFunds.Reader();

            List<string> list = new List<string>();
            list = reader.SupList;

            List<List<string>> list2 = new List<List<string>>();
            list2 = reader.FundList(@"C:\Temp\Marketing\Lists of Funds\Israel\Israeli funds.xlsx");
        }
    }
}
