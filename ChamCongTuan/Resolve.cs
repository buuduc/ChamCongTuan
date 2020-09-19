using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;


namespace ChamCongTuan
{
    class Resolve
    {
        List<Person> listPerson;
        List<String> listCongTac;
        ExcelWorksheet Worksheet;
        ExcelWorksheet Worksheet0;
        ExcelWorksheet Worksheet1;
        ExcelPackage excel = new ExcelPackage();
        public DateTime FinalDays;
        public DateTime FirstDays;
        public Resolve(List<Person> ListPerson, List<String> CongTacLst)
        {
            this.listPerson = ListPerson;
            this.listCongTac = CongTacLst;
            excel.Workbook.Worksheets.Add("Công Hành Chính");
            excel.Workbook.Worksheets.Add("Công tăng ca");

            this.Worksheet0 = excel.Workbook.Worksheets[0];
            this.Worksheet1 = excel.Workbook.Worksheets[1];
        }

        public void CreateNewFile(string @path)
        {
            
           
            FileInfo excelFile = new FileInfo(@path);
            excel.SaveAs(excelFile);

        }
        public void Process()
        {
            FinalDays = new DateTime(2020, 8, 31);

            Worksheet = Worksheet0;
            int i = 1;
            HeaderRow(i++);
            foreach (Person ps in listPerson)
            {
                RowData(i++, ps,ps.PubSalaryHours);
            }
            var range = Worksheet.Dimension;
            var FirstTableRange = Worksheet.Cells[Worksheet.Dimension.ToString()];
            FirstTableRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            FirstTableRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            FirstTableRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            FirstTableRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;


            Worksheet = Worksheet1;
            i = 1;
            HeaderRow(i++);
            foreach (Person ps in listPerson)
            {
                RowData(i++, ps, ps.OverSalaryHours);
            }
            range = Worksheet.Dimension;
            FirstTableRange = Worksheet.Cells[Worksheet.Dimension.ToString()];
            FirstTableRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            FirstTableRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            FirstTableRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            FirstTableRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;



        }
        private void Test()
        {
            var ps = this.listPerson.Find(x => x.MaNhanVien == "DH0225");
            ps.TinhCong();
            var date = new DateTime(2020, 8, 15);
     
            if (ps.CheckChamCong[date].ToString() == "x")
            {
                MessageBox.Show("dcm nha no");
            }

        }
        public void RowData(int row,Person ps, System.Collections.Hashtable list)
        {
            
            int j = 1;
            Worksheet.Column(j).AutoFit();
            Worksheet.Cells[row, j++].Value = ps.MaNhanVien;
            Worksheet.Column(j).AutoFit();
            Worksheet.Cells[row, j++].Value = ps.HoTen;
            Worksheet.Column(j).AutoFit();
            Worksheet.Cells[row, j++].Value = ps.NgaySinh;
            Worksheet.Column(j).AutoFit();
            Worksheet.Cells[row, j++].Value = ps.PhongBan;
            for (DateTime date = new DateTime(2020,8,1); FinalDays.CompareTo(date) >=0; date = date.AddDays(1.0))
            {
                ps.TinhCong();

                if (date.DayOfWeek == DayOfWeek.Sunday)
                {
                    Worksheet.Cells[row, j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Worksheet.Cells[row, j].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                }
                else if (ps.CheckChamCong[date] == null )
                {
                    Worksheet.Cells[row, j].Style.Fill.PatternType = ExcelFillStyle.DarkGray;
                    Worksheet.Cells[row, j].Style.Fill.BackgroundColor.SetColor(Color.Gray);
                }
                else if( ps.CheckChamCong[date].ToString()== "x" || ps.CheckChamCong[date].ToString() == "KL" )
                {
                    Worksheet.Cells[row, j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Worksheet.Cells[row, j].Style.Fill.BackgroundColor.SetColor(Color.Red);
                }
                else if (ps.CheckChamCong[date].ToString() == "Ô" || ps.CheckChamCong[date].ToString() == "P" || ps.CheckChamCong[date].ToString() == "P/2" || ps.CheckChamCong[date].ToString() == "Ô/2")
                {
                    Worksheet.Cells[row, j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Worksheet.Cells[row, j].Style.Fill.BackgroundColor.SetColor(Color.Aqua);
                }
                Worksheet.Cells[row, j++].Value = list[date];

            }
        


        }
        private void HeaderRow(int row)
        {
            int j = 1;
            Worksheet.Cells[row, j++].Value = "Mã Nhân Viên";
            Worksheet.Cells[row, j++].Value = "Họ Tên";
            Worksheet.Cells[row, j++].Value = "Ngày Sinh";
            Worksheet.Cells[row, j++].Value = "Phòng ban";
            for (DateTime date = new DateTime(2020, 8, 1); FinalDays.CompareTo(date) >= 0; date = date.AddDays(1.0))
            {
                string strday = date.DayOfWeek.ToString()+ "\n"+ date.Day.ToString() + "/" + date.Month.ToString()+"/"+ date.Year.ToString();
                Worksheet.Cells[row,j].Style.WrapText = true;
                Worksheet.Column(j).Width = 11.5;
                if (date.DayOfWeek == DayOfWeek.Sunday)
                {
                    Worksheet.Cells[row, j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Worksheet.Cells[row, j].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                }
                Worksheet.Cells[row, j++].Value = strday;
                //Worksheet.Cells[1, j++].Value = ps.PubSalaryHours[new DateTime(2020, 8, date.Day)];

            }
        }
    }
}
