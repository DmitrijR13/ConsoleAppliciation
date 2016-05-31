using ClosedXML.Excel;
using ConsoleApplication1.Database;
using ConsoleApplication9;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1.mainCode
{
    class InsertPeople
    {
        public void Ins1()
        {
            var wb2 = new XLWorkbook(@"C:\temp\Сведения о квартирах.xlsx");
            InsertPeopleDb insertPeopleDb = new InsertPeopleDb();
            for (int i = 2; i <= 256; i++)
            {
                try
                {
                    string str = insertPeopleDb.InsertPeople5("9100451",
                        Convert.ToInt32(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim()),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim(),
                        "0",
                        "Да",
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim());
                    if (str != "ЗАГРУЖЕНО")
                        Console.WriteLine(str);
                }
                catch
                {

                }

            }
            wb2.Save();
        }

        public void Ins2()
        {
            InsertPeopleDb insertPeopleDb = new InsertPeopleDb();
            DataTable dtHouse = new DataTable();
            DataTable dtPeople = new DataTable();
            var wb2 = new XLWorkbook(@"C:\temp\Копия Реестр исходных данных кап.ремонт.xlsx");
            string[] stringSeparators = new string[] { ", кв." };
            for (int i = 8; i <= 111; i++)
            {
                try
                {
                    string str = insertPeopleDb.InsertPeople("8800162",
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(8).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(10).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(9).Value).Trim());
                    if (str != "ЗАГРУЖЕНО")
                        Console.WriteLine(str);
                }
                catch
                {

                }

            }
            wb2.Save();
        }

        public void Ins3()
        {
            InsertPeopleDb insertPeopleDb = new InsertPeopleDb();
            DataTable dtHouse = new DataTable();
            DataTable dtPeople = new DataTable();
            var wb2 = new XLWorkbook(@"C:\temp\Копия Площадь 76.xlsx");
            for (int i = 3; i <= 441; i++)
            {
                try
                {
                    string str = insertPeopleDb.InsertPeople("9700035",
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim().Split('-')[1],
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim().Substring(4), "0",
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value), "0");
                    if (str != "ЗАГРУЖЕНО")
                        Console.WriteLine(str);
                }
                catch
                {

                }

            }
            wb2.Save();
        }

        public void Ins4()
        {
            InsertPeopleDb insertPeopleDb = new InsertPeopleDb();
            DataTable dtHouse = new DataTable();
            DataTable dtPeople = new DataTable();
            var wb2 = new XLWorkbook(@"C:\temp\Копия Фрунз. 8 В.xlsx");
            for (int i = 6; i <= 17; i++)
            {
                try
                {
                    string str = insertPeopleDb.InsertPeople("9700760",
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value), "30");
                    if (str != "ЗАГРУЖЕНО")
                        Console.WriteLine(str);
                }
                catch
                {

                }

            }
            wb2.Save();
        }
    }
}
