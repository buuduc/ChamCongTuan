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
        private int firstRow;
        List<Person> listPerson;
        List<String> listCongTac;
        ExcelWorksheet Worksheet;
        public ExcelWorksheet Worksheet0;
        public ExcelWorksheet Worksheet1;
        ExcelPackage excel = new ExcelPackage();
        public DateTime FinalDays;
        public DateTime FirstDays;
        public Resolve(List<Person> ListPerson, List<String> CongTacLst)
        {
            this.listPerson = ListPerson;
            this.listCongTac = CongTacLst;
            excel.Workbook.Worksheets.Add("Công Hành Chính");
            excel.Workbook.Worksheets.Add("Công Tăng Ca");

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
            firstRow = 1;
            Worksheet = Worksheet0;
            int i = firstRow;
            HeaderRowHanhChinh(i++,1);
            i++;
            foreach (Person ps in listPerson)
            {
                RowDataHanhChinh(i++, 1,ps,ps.PubSalaryHours);
            }
            var range = Worksheet.Dimension;
            var FirstTableRange = Worksheet.Cells[Worksheet.Dimension.ToString()];
            FirstTableRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            FirstTableRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            FirstTableRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            FirstTableRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;


            Worksheet = Worksheet1;
            i = firstRow;
            HeaderRowTangCa(i++,1);
            i+=2;
            foreach (Person ps in listPerson)
            {
                RowDataTangCa(i++,1, ps, ps.OverSalaryHours);
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
        public void RowDataHanhChinh(int row,int j,Person ps, System.Collections.Hashtable list)
        {
            
            
            Worksheet.Column(j).AutoFit();
            Worksheet.Cells[row, j++].Value = ps.PhongBan;
            Worksheet.Column(j).AutoFit();
            Worksheet.Cells[row, j++].Value = ps.MaNhanVien;
            Worksheet.Column(j).AutoFit();
            Worksheet.Cells[row, j++].Value = ps.HoTen;
            Worksheet.Column(j).AutoFit();
            Worksheet.Cells[row, j++].Value = ps.NgaySinh;
            Worksheet.Column(j).AutoFit();
            Worksheet.Cells[row, j++].Value = ps.MaNhanSu;
            var firstAddress = Worksheet.Cells[row, j].Address;
            for (DateTime date = FirstDays; FinalDays.CompareTo(date) >=0; date = date.AddDays(1.0))
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
                    Worksheet.Cells[row, j].Value = "KL";
                }
                else if (ps.CheckChamCong[date].ToString() == "Ô" || ps.CheckChamCong[date].ToString() == "P" || ps.CheckChamCong[date].ToString() == "P/2" || ps.CheckChamCong[date].ToString() == "Ô/2")
                {
                    Worksheet.Cells[row, j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Worksheet.Cells[row, j].Style.Fill.BackgroundColor.SetColor(Color.Aqua);
                }
                Worksheet.Cells[row, j++].Value = list[date];
                
            }
            var lastAddress = Worksheet.Cells[row, j-1].Address;
            Worksheet.Cells[row, j].Value = ps.HopDong;
            string HopDong = Worksheet.Cells[row, j].Address;
            // Ngày làm việc
            Worksheet.Cells[row, ++j].Formula="SUM("+ firstAddress+":"+lastAddress+")";
            string ngaylamviec = Worksheet.Cells[row, j].Address;
            //ngày nghỉ phép
            Worksheet.Cells[row, ++j].Formula =
                $"IF({HopDong} = \"CT\", COUNTIF({firstAddress}:{lastAddress}, \"P\")+SUMIF({firstAddress}:{lastAddress},0.5),0)";
            string ngaynghiphep = Worksheet.Cells[row, j].Address;
            //ngày nghỉ không lương
            Worksheet.Cells[row, ++j].Formula =
                $"IF(OR({HopDong} = \"TV\", {HopDong} = \"TT\"), COUNTIF({firstAddress}: {lastAddress}, \"KL\") + SUMIF({firstAddress}: {lastAddress}, 0.5), COUNTIF({firstAddress}: {lastAddress}, \"KL\"))";
            string ngaynghikhongluong = Worksheet.Cells[row, j].Address;
            //ngày nghỉ lễ
            Worksheet.Cells[row, ++j].Formula = $"COUNTIF({firstAddress}: {lastAddress},\"L\")";
            string ngaynghile = Worksheet.Cells[row, j].Address;
            // ngày được tính lương
            Worksheet.Cells[row, ++j].Formula = $" IF(OR({HopDong} = \"TT\", {HopDong} = \"TV\"), {ngaylamviec} + {ngaynghile}, {ngaylamviec} + {ngaynghiphep} + {ngaynghile})";
            string ngayduoctinhluong = Worksheet.Cells[row, j].Address;
            // ngày phep tinh den thang nay
            ++j;
            string ngaypheptinhdenthangnay = Worksheet.Cells[row, j].Address;
            // Ngày phép còn lại 
            Worksheet.Cells[row, ++j].Formula = $"{ngaynghiphep}-{ngaypheptinhdenthangnay}";
            string ngayphepconlai = Worksheet.Cells[row, j].Address;
            // Ngày tính lương thực tế
            Worksheet.Cells[row, ++j].Formula = $" IF({ngayphepconlai} < 0, {ngayduoctinhluong} - {ngaynghiphep} + {ngaypheptinhdenthangnay}, {ngayduoctinhluong})";
        }
        private void HeaderRowHanhChinh(int row,int j)
        {
            
            Worksheet.Cells[row, j, row + 1, j].Merge = true;
            Worksheet.Cells[row, j].Style.VerticalAlignment= ExcelVerticalAlignment.Center;
            Worksheet.Cells[row, j++].Value = "Phòng ban";
            Worksheet.Cells[row, j, row + 1, j].Merge = true;
            Worksheet.Cells[row, j].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            Worksheet.Cells[row, j++].Value = "Mã Nhân Viên";
            Worksheet.Cells[row, j, row + 1, j].Merge = true;
            Worksheet.Cells[row, j].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            Worksheet.Cells[row, j++].Value = "Họ Tên";
            Worksheet.Cells[row, j, row + 1, j].Merge = true;
            Worksheet.Cells[row, j].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            Worksheet.Cells[row, j++].Value = "Ngày Sinh";
            Worksheet.Cells[row, j, row + 1, j].Merge = true;
            Worksheet.Cells[row, j].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            Worksheet.Cells[row, j++].Value = "Chức Vụ";

            //var firstAddress = Worksheet.Cells[row, j++].Address;
            for (DateTime date = FirstDays; FinalDays.CompareTo(date) >= 0; date = date.AddDays(1.0))
            {
                string strday = date.DayOfWeek.ToString()+ "\n"+ date.Day.ToString() + "/" + date.Month.ToString()+"/"+ date.Year.ToString();
                Worksheet.Cells[row,j].Style.WrapText = true;
                Worksheet.Column(j).Width = 5;
                if (date.DayOfWeek == DayOfWeek.Sunday)
                {
                    Worksheet.Cells[row, j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Worksheet.Cells[row, j].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    Worksheet.Cells[row+1, j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Worksheet.Cells[row+1, j].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                }
                string a;
                
                Worksheet.Cells[row, j].Value = date.Day.ToString();
                Worksheet.Cells[row + 1, j++].Value = ConvertDayOfWeeks(date.DayOfWeek);

            }

            {
                Worksheet.Cells[row, j, row + 1, j].Merge = true;
                Worksheet.Cells[row, j].Value = "Hình thức hợp đồng";
                Worksheet.Cells[row, j].Style.WrapText = true;
                Worksheet.Cells[row, j].Style.Font.Bold = true;
                Worksheet.Cells[row, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Worksheet.Column(j).Width = 10;
                Worksheet.Column(j).Style.Font.Bold = true;
            }

            {
                Worksheet.Cells[row, ++j].Value = "Ngày làm việc";
                Worksheet.Cells[row, j, row + 1, j].Merge = true;
                Worksheet.Cells[row, j].Style.WrapText = true;
                Worksheet.Cells[row, j].Style.Font.Bold =true;
                Worksheet.Cells[row, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Worksheet.Column(j).Width = 10;
                Worksheet.Column(j).Style.Font.Bold = true;
            }
            {
                Worksheet.Cells[row, ++j].Value = "Ngày nghỉ phép";
                Worksheet.Cells[row, j, row + 1, j].Merge = true;
                Worksheet.Cells[row, j].Style.WrapText = true;
                Worksheet.Cells[row, j].Style.Font.Bold = true;
                Worksheet.Cells[row, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Worksheet.Column(j).Width = 10;
                Worksheet.Column(j).Style.Font.Bold = true;
            }
            {
                Worksheet.Cells[row, ++j].Value = "Ngày nghỉ không lương";
                Worksheet.Cells[row, j, row + 1, j].Merge = true;
                Worksheet.Cells[row, j].Style.WrapText = true;
                Worksheet.Cells[row, j].Style.Font.Bold = true;
                Worksheet.Cells[row, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Worksheet.Column(j).Width = 10;
                Worksheet.Column(j).Style.Font.Bold = true;
            }
            {
                Worksheet.Cells[row, ++j].Value = "Ngày nghỉ lễ";
                Worksheet.Cells[row, j, row + 1, j].Merge = true;
                Worksheet.Cells[row, j].Style.WrapText = true;
                Worksheet.Cells[row, j].Style.Font.Bold = true;
                Worksheet.Cells[row, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Worksheet.Column(j).Width = 10;
                Worksheet.Column(j).Style.Font.Bold = true;
            }
            {
                Worksheet.Cells[row, ++j].Value = "Ngày được tính lương";
                Worksheet.Cells[row, j, row + 1, j].Merge = true;
                Worksheet.Cells[row, j].Style.WrapText = true;
                Worksheet.Cells[row, j].Style.Font.Bold = true;
                Worksheet.Cells[row, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Worksheet.Column(j).Width = 10;
                Worksheet.Column(j).Style.Font.Bold = true;
            }
            {
                Worksheet.Cells[row, ++j].Value = "Ngày phép năm tính đến tháng này";
                Worksheet.Cells[row, j, row + 1, j].Merge = true;
                Worksheet.Cells[row, j].Style.WrapText = true;
                Worksheet.Cells[row, j].Style.Font.Bold = true;
                Worksheet.Cells[row, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Worksheet.Column(j).Width = 10;
                Worksheet.Column(j).Style.Font.Bold = true;
            }
            {
                Worksheet.Cells[row, ++j].Value = "Ngày phép còn lại";
                Worksheet.Cells[row, j, row + 1, j].Merge = true;
                Worksheet.Cells[row, j].Style.WrapText = true;
                Worksheet.Cells[row, j].Style.Font.Bold = true;
                Worksheet.Cells[row, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Worksheet.Column(j).Width = 10;
                Worksheet.Column(j).Style.Font.Bold = true;
            }
            {
                Worksheet.Cells[row, ++j].Value = "Ngày tính lương thực tế";
                Worksheet.Cells[row, j, row + 1, j].Merge = true;
                Worksheet.Cells[row, j].Style.WrapText = true;
                Worksheet.Cells[row, j].Style.Font.Bold = true;
                Worksheet.Cells[row, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Worksheet.Column(j).Width = 10;
                Worksheet.Column(j).Style.Font.Bold = true;
            }

        }
        private void HeaderRowTangCa(int row,int j)
        {
            Worksheet.Cells[row, j, row + 2, j].Merge = true;
            Worksheet.Cells[row, j].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            Worksheet.Cells[row, j++].Value = "Phòng ban";
            Worksheet.Cells[row, j, row + 2, j].Merge = true;
            Worksheet.Cells[row, j].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            Worksheet.Cells[row, j++].Value = "Mã Nhân Viên";
            Worksheet.Cells[row, j, row + 2, j].Merge = true;
            Worksheet.Cells[row, j].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            Worksheet.Cells[row, j++].Value = "Họ Tên";
            Worksheet.Cells[row, j, row + 2, j].Merge = true;
            Worksheet.Cells[row, j].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            Worksheet.Cells[row, j++].Value = "Ngày Sinh";
            Worksheet.Cells[row, j, row + 2, j].Merge = true;
            Worksheet.Cells[row, j].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            Worksheet.Cells[row, j++].Value = "Chức vụ";
            var firstcell = Worksheet.Cells[row, j];
            row += 1;
            //var firstAddress = Worksheet.Cells[row, j++].Address;
            for (DateTime date = FirstDays; FinalDays.CompareTo(date) >= 0; date = date.AddDays(1.0))
            {
                string strday = date.DayOfWeek.ToString() + "\n" + date.Day.ToString() + "/" + date.Month.ToString() + "/" + date.Year.ToString();
                Worksheet.Cells[row, j].Style.WrapText = true;
                Worksheet.Column(j).Width = 5;
                if (date.DayOfWeek == DayOfWeek.Sunday)
                {
                    Worksheet.Cells[row, j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Worksheet.Cells[row, j].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                    Worksheet.Cells[row + 1, j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Worksheet.Cells[row + 1, j].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                }

                Worksheet.Cells[row, j].Value = date.Day.ToString();
                Worksheet.Cells[row + 1, j++].Value = ConvertDayOfWeeks(date.DayOfWeek);

            }
            var latsCell = Worksheet.Cells[row-1, j-1];
            Worksheet.Cells[$"{firstcell.Address}:{latsCell.Address}"].Merge = true; //a1 den aj1
            
            {
                Worksheet.Cells[row, j, row + 1, j].Merge = true;
                Worksheet.Cells[row, j].Value = "Ngày CN (8h/ngày)";
                Worksheet.Cells[row, j].Style.WrapText = true;
                Worksheet.Cells[row, j].Style.Font.Bold = true;
                Worksheet.Cells[row, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Worksheet.Column(j).Width = 10;
                Worksheet.Column(j).Style.Font.Bold = true;
            }

            {
                Worksheet.Cells[row, ++j].Value = "Tăng ca còn lại";
                Worksheet.Cells[row, j, row + 1, j].Merge = true;
                Worksheet.Cells[row, j].Style.WrapText = true;
                Worksheet.Cells[row, j].Style.Font.Bold = true;
                Worksheet.Cells[row, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Worksheet.Column(j).Width = 10;
                Worksheet.Column(j).Style.Font.Bold = true;
            }
            Worksheet.Cells[row-1, j-1, row-1, j].Merge = true;
            Worksheet.Cells[row - 1, j - 1].Value= "Ngày tăng ca tính cơm";
            Worksheet.Cells[row - 1, j - 1].Style.WrapText = true;
            {
                j++;
                Worksheet.Cells[row-1, j, row + 1, j].Merge = true;
                Worksheet.Cells[row-1, j].Value = "TC (NGÀY \n THƯỜNG 1.5) GIỜ";
                Worksheet.Cells[row, j].Style.WrapText = true;
                Worksheet.Cells[row, j].Style.Font.Bold = true;
                Worksheet.Cells[row, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Worksheet.Column(j).Width = 10;
                Worksheet.Column(j).Style.Font.Bold = true;
            }
            {
                j++;
                Worksheet.Cells[row - 1, j, row + 1, j].Merge = true;
                Worksheet.Cells[row-1, j].Value = "TC (NGÀY \n CN 2.0) GIỜ";
                Worksheet.Cells[row, j].Style.WrapText = true;
                Worksheet.Cells[row, j].Style.Font.Bold = true;
                Worksheet.Cells[row, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Worksheet.Column(j).Width = 10;
                Worksheet.Column(j).Style.Font.Bold = true;
            }
            {
                j++;
                Worksheet.Cells[row - 1, j, row + 1, j].Merge = true;
                Worksheet.Cells[row-1, j].Value = "TC NGÀY \n LỄ, TẾT (3.0 GIỜ)";
                Worksheet.Cells[row, j].Style.WrapText = true;
                Worksheet.Cells[row, j].Style.Font.Bold = true;
                Worksheet.Cells[row, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Worksheet.Column(j).Width = 10;
                Worksheet.Column(j).Style.Font.Bold = true;
            }
            {
                j++;
                Worksheet.Cells[row - 1, j, row + 1, j].Merge = true;
                Worksheet.Cells[row-1, j].Value = "TỔNG GIỜ TĂNG CA";
                Worksheet.Cells[row, j].Style.WrapText = true;
                Worksheet.Cells[row, j].Style.Font.Bold = true;
                Worksheet.Cells[row, j].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                Worksheet.Column(j).Width = 10;
                Worksheet.Column(j).Style.Font.Bold = true;
            }


        }
        public void RowDataTangCa(int row, int j, Person ps, System.Collections.Hashtable list)
        {
            
          
            Worksheet.Column(j).AutoFit();
            Worksheet.Cells[row, j++].Value = ps.PhongBan;
            Worksheet.Column(j).AutoFit();
            Worksheet.Cells[row, j++].Value = ps.MaNhanVien;
            Worksheet.Column(j).AutoFit();
            Worksheet.Cells[row, j++].Value = ps.HoTen;
            Worksheet.Column(j).AutoFit();
            Worksheet.Cells[row, j++].Value = ps.NgaySinh;
            Worksheet.Column(j).AutoFit();
            Worksheet.Cells[row, j++].Value = ps.MaNhanSu;
            var firstAddress = Worksheet.Cells[row, j].Address;
            var firstcolumn = j;
            for (DateTime date = FirstDays; FinalDays.CompareTo(date) >= 0; date = date.AddDays(1.0))
            {
                ps.TinhCong();

                if (date.DayOfWeek == DayOfWeek.Sunday)
                {
                    Worksheet.Cells[row, j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Worksheet.Cells[row, j].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                }
                else if (ps.CheckChamCong[date] == null)
                {
                    Worksheet.Cells[row, j].Style.Fill.PatternType = ExcelFillStyle.DarkGray;
                    Worksheet.Cells[row, j].Style.Fill.BackgroundColor.SetColor(Color.Gray);
                }
                else if (ps.CheckChamCong[date].ToString() == "x" || ps.CheckChamCong[date].ToString() == "KL")
                {
                    Worksheet.Cells[row, j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Worksheet.Cells[row, j].Style.Fill.BackgroundColor.SetColor(Color.Red);
                    Worksheet.Cells[row, j].Value = "KL";
                }
                else if (ps.CheckChamCong[date].ToString() == "Ô" || ps.CheckChamCong[date].ToString() == "P" || ps.CheckChamCong[date].ToString() == "P/2" || ps.CheckChamCong[date].ToString() == "Ô/2")
                {
                    Worksheet.Cells[row, j].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    Worksheet.Cells[row, j].Style.Fill.BackgroundColor.SetColor(Color.Aqua);
                }
                Worksheet.Cells[row, j++].Value = list[date];

            }

            var lastColumn = j - 1;
            var lastAddress = Worksheet.Cells[row, lastColumn].Address;
            j -= 1;
            //ngay chu nhat
            Worksheet.Cells[row, ++j].Formula = $"COUNTIF({firstAddress}:{lastAddress},\" >= 8\")";
            string ngaychunhat= Worksheet.Cells[row, j].Address;
            //Tăng ca còn lại
            Worksheet.Cells[row, ++j].Formula = $"COUNTIF({firstAddress}:{lastAddress},\" >= 2\")-{ngaychunhat}";
            string tangcaconlai = Worksheet.Cells[row, j].Address;
            //Tang ca 1.5 gio
            
            string tangca15 = Worksheet.Cells[row, ++j].Address;

            // Tang ca 2.0 gio
            Worksheet.Cells[row,++ j].Formula = $"=(SUMIF({Worksheet.Cells[firstRow+2,firstcolumn,firstRow+2,lastColumn].Address},\"CN\",{firstAddress}:{lastAddress}))";
            string tangca20= Worksheet.Cells[row, j].Address;

            // tang ca 3.0 gio
            string tangca30 = Worksheet.Cells[row, ++j].Address;
            Worksheet.Cells[tangca15].Formula = $"SUM({firstAddress}:{lastAddress}) - {tangca20} - {tangca30}";

            // TỔNG GIỜ TĂNG CA
            Worksheet.Cells[row, ++j].Formula= $"SUM({tangca15}:{tangca30})";

        }

        private void FormatCondition()
        {
            // var range = Worksheet.Cells[1,10,1,10]; //="B2:B11"
            // //find top 3 items, depend on the bigger value
            // var _cond4 = Worksheet.ConditionalFormatting.AddExpression(_formatRangeAddress);
            // //or find top 30% items, depend on the bigger value
            // //var rule2 = range.AddConditionalFormat().IsTop(30, TopBottomType.Percent);
            // rule2.Style.Fill
        }
        private string ConvertDayOfWeeks(System.DayOfWeek day)
        {
            string a;
            switch (day)
            {
                case DayOfWeek.Sunday:
                    return "CN";           
                case DayOfWeek.Monday:
                    return "T2";
                case DayOfWeek.Tuesday:
                    return "T3";
                case DayOfWeek.Wednesday:
                    return "T4";
                case DayOfWeek.Thursday:
                    return "T5";
                case DayOfWeek.Friday:
                    return "T6";
                case DayOfWeek.Saturday:
                    return "T7";
                   
                default:
                    return "";
            }
            
        }
    }
}
