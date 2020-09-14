using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ChamCongTuan
{
    class Person
    {
        public String MaNhanVien { get; set; }
        public String HoTen { get; set; }
        public String NgaySinh { get; set; }
        public String PhongBan { get; set; }
        public String ViTri { get; set; }
        public String SDT { get; set; }
        public double CongHanhChinh()
        {
            double pubSalary = 0;
            foreach (DictionaryEntry C in this.ChamCong)
            {
                var day = C.Key.ToString();
                DateTime date = new DateTime(2020, Int32.Parse(day.Substring(3, 2)), Int32.Parse(day.Substring(0, 2)));
                if (double.TryParse(C.Value.ToString(), out double number))
                {
                    if (number == 1)
                    {
                        pubSalary++;
                    }
                }

            }
            return pubSalary;
        }
        public Hashtable ChamCong = new Hashtable();


    }
}
