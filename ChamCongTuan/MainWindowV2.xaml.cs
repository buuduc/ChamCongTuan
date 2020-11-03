using Microsoft.Win32;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections;
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
using ChamCongTuanV2.Properties;

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
            //test();
        }
        private List<Person> ListPerson = new List<Person>();
        private List<string> CongTaclst = new List<string>();
        private ExcelWorksheet Worksheet;
        private string HoSoPath;
        private string DuLieuPath;
        private List<String> listDay = new List<string>();
        private Hashtable MaNhanSu = new Hashtable();
        private void NhapHoSoBtn(object sender, RoutedEventArgs e)
        {

            try
            {
                NhapHoSoCommand();
                MessageBox.Show("Nhập hồ sơ thành công !", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.ToString(), "Error!",MessageBoxButton.OK,MessageBoxImage.Error);
            }

      

        }
        private void NhapHoSoCommand()
        {
        ExcelWorksheet workSheet;
       
            using (ExcelPackage MaNS = new ExcelPackage(new FileInfo(Settings.Default.MaNhanSuPath)))
            {
                // lấy ra sheet đầu tiên để thao tác
                workSheet = MaNS.Workbook.Worksheets[0];
                for (int i = workSheet.Dimension.Start.Row+1; i <= workSheet.Dimension.End.Row; i++)
                {
                    int j = 1;
                    string key = workSheet.Cells[i, j++].Value.ToString();
                    string ob= workSheet.Cells[i, j++].Value.ToString();
                    MaNhanSu.Add(key, ob);
                }
            }

            
            // mở file excel
            var package = new ExcelPackage(new FileInfo(this.HoSoPath));

            // lấy ra sheet đầu tiên để thao tác
            workSheet = package.Workbook.Worksheets[0];

            for (int i = workSheet.Dimension.Start.Row + 1; i <= workSheet.Dimension.End.Row; i++)
            {
                try
                {

                    // biến j biểu thị cho một column trong file
                    int j = 1;
                    int headerrow = 1;
                    Person person = new Person();
                    while (j<=workSheet.Dimension.End.Column)
                        person.DataList.Add(workSheet.Cells[headerrow, j].Value.ToString(), workSheet.Cells[i, j++].Value.ToString());
                    

                    //person.MaNhanVien = workSheet.Cells[i, j++].Value.ToString();
                    //person.HoTen = workSheet.Cells[i, j++].Value.ToString();
                    //CongTaclst.Add(workSheet.Cells[i, j].Value.ToString());
                    //person.PhongBan = workSheet.Cells[i, j++].Value.ToString();
                    //person.ViTri = workSheet.Cells[i, j++].Value.ToString();
                    //person.NgaySinh = workSheet.Cells[i, j++].Value.ToString();
                    //person.SDT = workSheet.Cells[i, j++].Value.ToString();
                    //object a = person.ViTri;
                    person.MaNhanSu = MaNhanSu[person.ViTri].ToString();
                    ListPerson.Add(person);
                }
                catch (Exception exe)
                {

                }
            }
            Datagrid.ItemsSource = ListPerson;
            CongTaclst = CongTaclst.Distinct().ToList();
        }

        private void NhapDuLieuBtn(object sender, RoutedEventArgs e)
        {
            //List<Person> PersonList = new List<Person>();
            try
            {
               NhapDuLieuCommand();
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.ToString(), "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }
        private void NhapDuLieuCommand()
        {
            // mở file excel
            var package = new ExcelPackage(new FileInfo(this.DuLieuPath));

            // lấy ra sheet đầu tiên để thao tác
            ExcelWorksheet workSheet = package.Workbook.Worksheets[0];
            //workSheet.Cells[4, 1].Value.ToString();
            var check = true;
            for (int i = workSheet.Dimension.Start.Row + 2; i <= workSheet.Dimension.End.Row; i++)
            {
                Person person = this.ListPerson.Find(ps => ps.MaNhanVien == workSheet.Cells[i, 1].Value.ToString());
                try
                {
                    //for (int index = 1; index <= 31; index++)
                    int index = 1;
                  
                    while (workSheet.Cells[2, index + 4].Value.ToString()[2] == '/')
                    {
                        if (check)
                        {
                            listDay.Add(workSheet.Cells[2, index + 4].Value.ToString());
                            
                        }
                        if (workSheet.Cells[i, index + 4].Value != null)
                        {
                            person.ChamCong.Add(workSheet.Cells[2, index + 4].Value, workSheet.Cells[i, index + 4].Value);

                        }
                        index++;
                    }
                    check = false;

                }
                catch (Exception exe)
                {

                }
            }
            //
            Datagrid.ItemsSource = this.ListPerson;
            FirstDayBox.ItemsSource = listDay;
            LastDayBox.ItemsSource = listDay;
            Yearbox.Text = DateTime.Now.Year.ToString();
            MessageBox.Show("Nhập dữ liệu thành công !", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);

        }


        private void ExportDataBtn(object sender, RoutedEventArgs e)
        {
           ExportDataCommand();
            try
            {
                
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.ToString(), "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }
        private void ExportDataCommand()
        {
            Resolve tt = new Resolve(ListPerson, CongTaclst);
            tt.FirstDays= new DateTime(Int32.Parse(Yearbox.Text.ToString()), Int32.Parse(FirstDayBox.Text.ToString().Substring(3, 2)), Int32.Parse(FirstDayBox.Text.ToString().Substring(0, 2)));
            tt.FinalDays = new DateTime(Int32.Parse(Yearbox.Text.ToString()), Int32.Parse(LastDayBox.Text.ToString().Substring(3, 2)), Int32.Parse(LastDayBox.Text.ToString().Substring(0, 2)));
            tt.Process();
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            saveFileDialog.FilterIndex = 1;
            saveFileDialog.FileName =@"BCC ngày " + @FirstDayBox.Text.Substring(0,2) + " - " + @LastDayBox.Text.Substring(0, 2) + " tháng "+ LastDayBox.Text.ToString().Substring(3, 2) +" năm " + Yearbox.Text + ".xlsx";
            if (saveFileDialog.ShowDialog() == true)
            {
                string path = saveFileDialog.FileName;
                tt.CreateNewFile(@path);
                MessageBox.Show("Xuất file thành công !", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                
            }

        }
       private void test()
        {
            this.HoSoPath = @"E:\OneDrive - poxz\User\ADMIN\Desktop\Ho-so-nhan-su.xlsx";
            this.DuLieuPath = @"E:\OneDrive - poxz\User\ADMIN\Documents\GitHub\ChamCongTuanV2\ChamCongTuan\Bang-cham-cong-[8-2020].xlsx";
            NhapHoSoCommand();
            NhapDuLieuCommand();
            Resolve tt = new Resolve(ListPerson, CongTaclst);
            tt.FirstDays = new DateTime(2020, 08, 01);
            tt.FinalDays = new DateTime(2020, 08, 31);
            tt.Process();
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            saveFileDialog.FilterIndex = 1;
            //saveFileDialog.FileName = @"BCC ngày " + @FirstDayBox.Text.Substring(0, 2) + " - " + @LastDayBox.Text.Substring(0, 2) + " tháng " + LastDayBox.Text.ToString().Substring(3, 2) + " năm " + Yearbox.Text + ".xlsx";
            if (saveFileDialog.ShowDialog() == true)
            {
                string path = saveFileDialog.FileName;
                tt.CreateNewFile(@path);
                MessageBox.Show("Xuất file thành công !", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);

            }
            //this.NhapHoSo.Click += new RoutedEventHandler(NhapDuLieuBtn);
            //this.NhapDuLieu.Click += new RoutedEventHandler(NhapHoSoBtn);
        }

        private void NhapHosoBrowseBtn(object sender, RoutedEventArgs e)
        {

            OpenFileDialog openFileDialog = new OpenFileDialog();

            //openFileDialog.InitialDirectory = "c:\\";
            openFileDialog.Filter = "excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == true)
            {
                //Get the path of specified file
                this.HoSoPath = @openFileDialog.FileName;
                NhapHoSoLabel.Foreground = Brushes.Green;
                NhapHoSoLabel.Content = "Đã chọn !";

            }
            
        }

  

        private void NhapDuLieuBrowseBtn(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            //openFileDialog.InitialDirectory = "c:\\";
            openFileDialog.Filter = "excel files (*.txt)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == true)
            {
                //Get the path of specified file
                this.DuLieuPath = @openFileDialog.FileName;
                NhapDuLieuLabel.Foreground = Brushes.Green;
                NhapDuLieuLabel.Content = "Đã chọn !";
            }
        }

       
    }
}


