using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;

namespace ChamCongTuan
{
    class Resolve
    {
        List<Person> listPerson;
        List<String> listCongTac;
        ExcelWorksheet Worksheet;
        ExcelPackage excel = new ExcelPackage();
        public Resolve(List<Person> ListPerson, List<String> CongTacLst)
        {
            this.listPerson = ListPerson;
            this.listCongTac = CongTacLst;
            excel.Workbook.Worksheets.Add("Worksheet1");
            Worksheet = excel.Workbook.Worksheets[0];
        }

        public void CreateNewFile(string @path)
        {
            
           
            FileInfo excelFile = new FileInfo(@path);
            excel.SaveAs(excelFile);

        }
        public void Process()
        {
            int i = 1;
            //Person ps = this.listPerson[59];


            //int j = 1;
            //Worksheet.Cells[i, j++].Value = "Mã Nhân Viên";
            //Worksheet.Cells[i, j++].Value = "Ho Tên";
            //Worksheet.Cells[i, j++].Value = "Ngày sinh";
            //Worksheet.Cells[i, j++].Value = "Phòng ban";
            //var a = new DateTime(2020, 8, 30);
            //for (DateTime date = new DateTime(2020, 8, 1); a.CompareTo(date) >= 0; date = date.AddDays(1.0))
            //{
            //    string strday = date.Day.ToString() + "/" + date.Month.ToString();
            //    Worksheet.Cells[i, j++].Value = strday;
            //    //Worksheet.Cells[1, j++].Value = ps.PubSalaryHours[new DateTime(2020, 8, date.Day)];

            //}



            HeaderRow(i++);
            foreach (Person ps in listPerson)
            {
                RowData(i++,ps);
            }



        }
        public void RowData(int row,Person ps)
        {
            
            int j = 1;
            
            Worksheet.Cells[row, j++].Value = ps.MaNhanVien;
            Worksheet.Cells[row, j++].Value = ps.HoTen;
            Worksheet.Cells[row, j++].Value = ps.NgaySinh;
            Worksheet.Cells[row, j++].Value = ps.PhongBan;
            var a = new DateTime(2020, 8, 30);
            for (DateTime date = new DateTime(2020,8,1); a.CompareTo(date) >=0; date = date.AddDays(1.0))
            {
                ps.TinhCong();
                Worksheet.Cells[row, j++].Value = ps.PubSalaryHours[new DateTime(2020, 8, date.Day)];

            }
        


        }
        private void HeaderRow(int row)
        {
            int j = 1;
            Worksheet.Cells[row, j++].Value = "Mã Nhân Viên";
            Worksheet.Cells[row, j++].Value = "Ho Tên";
            Worksheet.Cells[row, j++].Value = "Ngày sinh";
            Worksheet.Cells[row, j++].Value = "Phòng ban";
            var a = new DateTime(2020, 8, 30);
            for (DateTime date = new DateTime(2020, 8, 1); a.CompareTo(date) >= 0; date = date.AddDays(1.0))
            {
                string strday = date.DayOfWeek.ToString()+ "/n"+ date.Day.ToString() + "/" + date.Month.ToString();
                Worksheet.Cells[row, j++].Value = strday;
                //Worksheet.Cells[1, j++].Value = ps.PubSalaryHours[new DateTime(2020, 8, date.Day)];

            }
        }
    }
}
