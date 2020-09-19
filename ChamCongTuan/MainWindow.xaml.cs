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
            var package = new ExcelPackage(new FileInfo(@"E:\OneDrive - poxz\User\ADMIN\Documents\GitHub\ChamCongTuan\ChamCongTuan\HoSoNhanSu.xlsx"));

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
            var package = new ExcelPackage(new FileInfo(@"E:\OneDrive - poxz\User\ADMIN\Documents\GitHub\ChamCongTuan\ChamCongTuan\Bang-cham-cong-[8-2020].xlsx"));

            // lấy ra sheet đầu tiên để thao tác
            ExcelWorksheet workSheet = package.Workbook.Worksheets[0];
            //workSheet.Cells[4, 1].Value.ToString();

            for (int i = workSheet.Dimension.Start.Row + 2; i <= workSheet.Dimension.End.Row; i++)
            {
                Person person = this.ListPerson.Find(ps => ps.MaNhanVien == workSheet.Cells[i, 1].Value.ToString());
                try
                {
                    //for (int index = 1; index <= 31; index++)
                    int index = 1;
                    while (workSheet.Cells[2, index + 4].Value.ToString()[2] == '/')
                    {
                        if (workSheet.Cells[i, index + 4].Value != null)
                        {
                            person.ChamCong.Add(workSheet.Cells[2, index + 4].Value, workSheet.Cells[i, index + 4].Value);

                        }
                        index++;
                    }
                }
                catch (Exception exe)
                {

                }
            }
            //
            Datagrid.ItemsSource = this.ListPerson;
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
        private Stream CreateExcelFile(Stream stream = null)
        {
            
            using (var excelPackage = new ExcelPackage(stream ?? new MemoryStream()))
            {
                // Tạo author cho file Resolve
                excelPackage.Workbook.Properties.Author = "Hanker";
                // Tạo title cho file Resolve
                excelPackage.Workbook.Properties.Title = "EPP test background";
                // thêm tí comments vào làm màu 
                excelPackage.Workbook.Properties.Comments = "This is my fucking generated Comments";
                // Add Sheet vào file Resolve
                excelPackage.Workbook.Worksheets.Add("First Sheet");
                // Lấy Sheet bạn vừa mới tạo ra để thao tác 
                var workSheet = excelPackage.Workbook.Worksheets[0];
                // Đổ data vào Resolve file
                workSheet.Cells[1, 1].LoadFromCollection(ListPerson, true, TableStyles.Dark9);
                // BindingFormatForExcel(workSheet, list);
                excelPackage.Save();
                return excelPackage.Stream;
            }
        }
    }
}


