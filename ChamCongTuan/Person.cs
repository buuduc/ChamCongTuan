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
        
        public String MaNhanVien { get { return DataList["Mã NV"]==null?"":DataList["Mã NV"].ToString(); }}
        public String HoTen { get { return DataList["Họ và tên"] == null ? "" : DataList["Họ và tên"].ToString(); } }
        public String NgaySinh { get { return DataList["Ngày sinh"] == null ? "" : DataList["Ngày sinh"].ToString(); } }
        public String PhongBan { get { return DataList["Phòng ban"] == null ? "" : DataList["Phòng ban"].ToString(); } }
        public String ViTri { get { return DataList["Vị trí"] == null ? "" : DataList["Vị trí"].ToString(); } }
        public String SDT { get { return DataList["Điện thoại"] == null ? "" : DataList["Điện thoại"].ToString(); } }
        public String HopDong { get { return DataList["Tên hợp đồng"] == null ? "" : DataList["Tên hợp đồng"].ToString(); } }
        public String MaNhanSu { get; set; }
        public void TinhCong()
        {
            DeleteAll();
            double pubSalary = 0;
            
            double overSalary = 0;
            foreach (DictionaryEntry C in this.ChamCong)
            {
                var day = C.Key.ToString();
                DateTime date = new DateTime(2020, Int32.Parse(day.Substring(3, 2)), Int32.Parse(day.Substring(0, 2)));
                if (C.Value != null)
                {
                    CheckChamCong.Add(date, C.Value.ToString());

                }
                if (date.DayOfWeek == System.DayOfWeek.Sunday)
                {
                    if (double.TryParse(C.Value.ToString(), out double number))
                    {
                        double hoursTime = (number) * 10;
                        hoursTime = (Math.Floor((hoursTime - 2) * 2) / 2);
                        if (hoursTime >= 10)
                        {
                            
                            overSalary += 8;
                            OverSalaryHours.Add(date, 8);

                        }
                        else if (hoursTime > 8 & hoursTime < 10)
                        {
                            overSalary += hoursTime - 2;
                            OverSalaryHours.Add(date, hoursTime - 2);
                        }
                        else
                        {
                            overSalary += hoursTime;
                            OverSalaryHours.Add(date, hoursTime);
                        }
                        

                    }
                }
                else
                {
                    if (double.TryParse(C.Value.ToString(), out double number))
                    {
                        if (number >= 0.875)
                        {
                            pubSalary++;
                            overSalary += (Math.Floor(((number - 1) * 16) * 2) / 2);
                            if (number >= 1)
                            {
                                    PubSalaryHours.Add(date, 1);
                                    OverSalaryHours.Add(date, (Math.Floor(((number - 1) * 16) * 2) / 2));
                            }
                        }
                        else if(number < 0.875)
                        {
                            if (number >= 0.5)
                            {
                                pubSalary+=0.5;
                                PubSalaryHours.Add(date, 0.5);
                            }
                            if (number < 0.5)
                            {
                                pubSalary +=0;
                                PubSalaryHours.Add(date, 0);
                            }
                        }

                    }
                    else if (C.Value.ToString() == "Ô" | C.Value.ToString() == "P")
                    {
                        pubSalary++;
                        PubSalaryHours.Add(date, 1);
                    }
                    else if (C.Value.ToString() == "P/2" | C.Value.ToString() == "Ô/2")
                    {
                        pubSalary += 0.5;
                        PubSalaryHours.Add(date, 0.5);
                    }
                }
            }
            //return pubSalary;
        }
        public void DeleteAll()
        {
            PubSalaryHours.Clear();
            OverSalaryHours.Clear();
            CheckChamCong.Clear();
        }
       
        public Hashtable CheckChamCong = new Hashtable();
        public Hashtable ChamCong = new Hashtable();
        public Hashtable PubSalaryHours = new Hashtable();
        public Hashtable OverSalaryHours = new Hashtable();
        public Hashtable DataList =  new Hashtable();

    }
}
