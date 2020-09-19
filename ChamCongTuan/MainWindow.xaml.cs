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
        }
        private List<Person> ListPerson = new List<Person>();
        private List<string> CongTaclst = new List<string>();
        private ExcelWorksheet Worksheet;
        private string HoSoPath;
        private string DuLieuPath;
        private List<String> listDay = new List<string>();

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

            // mở file excel
            var package = new ExcelPackage(new FileInfo(this.HoSoPath));

            // lấy ra sheet đầu tiên để thao tác
            ExcelWorksheet workSheet = package.Workbook.Worksheets[0];

            // duyệt tuần tự từ dòng thứ 2 đến dòng cuối cùng của file. lưu ý file excel bắt đầu từ số 1 không phải số 0
            for (int i = workSheet.Dimension.Start.Row + 2; i <= workSheet.Dimension.End.Row; i++)
            {
                try
                {

                    // biến j biểu thị cho một column trong file
                    int j = 1;

                    // lấy ra cột họ tên tương ứng giá trị tại vị trí [i, 1]. i lần đầu là 2
                    // tăng j lên 1 đơn vị sau khi thực hiện xong câu lệnh

                    // lấy ra cột ngày sinh tương ứng giá trị tại vị trí [i, 2]. i lần đầu là 2
                    // tăng j lên 1 đơn vị sau khi thực hiện xong câu lệnh
                    // lấy ra giá trị ngày tháng và ép kiểu thành DateTime  
                    Person person = new Person();
                    person.MaNhanVien = workSheet.Cells[i, j++].Value.ToString();
                    person.HoTen = workSheet.Cells[i, j++].Value.ToString();
                    CongTaclst.Add(workSheet.Cells[i, j].Value.ToString());
                    person.PhongBan = workSheet.Cells[i, j++].Value.ToString();
                    person.ViTri = workSheet.Cells[i, j++].Value.ToString();
                    person.NgaySinh = workSheet.Cells[i, j++].Value.ToString();
                    person.SDT = workSheet.Cells[i, j++].Value.ToString();
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
                MessageBox.Show("Nhập dữ liệu thành công !", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
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
        }


        private void ExportDataBtn(object sender, RoutedEventArgs e)
        {

            try
            {
                ExportDataCommand();
                MessageBox.Show("Xuất file thành công !", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ee)
            {
                MessageBox.Show(ee.ToString(), "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }
        private void ExportDataCommand()
        {
            Resolve tt = new Resolve(ListPerson, CongTaclst);
            tt.Process();
            tt.CreateNewFile(@"E:\OneDrive - poxz\User\ADMIN\Desktop\Tesst\test.xlsx");

        }
       

        private void NhapHosoBrowseBtn(object sender, RoutedEventArgs e)
        {

            OpenFileDialog openFileDialog = new OpenFileDialog();

            //openFileDialog.InitialDirectory = "c:\\";
            openFileDialog.Filter = "excel files (*.txt)|*.xlsx|All files (*.*)|*.*";
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


