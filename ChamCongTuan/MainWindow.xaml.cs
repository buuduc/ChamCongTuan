using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
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

namespace ChamCongTuan
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.Commercial;

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var package = new ExcelPackage(new FileInfo(@"E:\OneDrive - poxz\User\ADMIN\Downloads\ExcelInCsharp-20200912T021252Z-001\ExcelInCsharp\ExcelInCsharp\bin\Debug\ImportData.xlsx"));
            ExcelWorksheet workSheet = package.Workbook.Worksheets[0];
        }
    }
}

