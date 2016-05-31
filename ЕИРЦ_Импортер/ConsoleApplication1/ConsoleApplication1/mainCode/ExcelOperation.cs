using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1.mainCode
{
    class ExcelOperation
    {
        public void Sum2File()
        {
            Dictionary<string, string> workBooks = new Dictionary<string, string>();
            workBooks.Add("Макс_итог_плюсом", "Макс (3)");
            workBooks.Add("Согласие_итог_плюсом", "Согласие");
            workBooks.Add("УралСиб_итог_плюсом(2)", "УралСиб (5)");
            foreach (KeyValuePair<string, string> books in workBooks)
            {
                Console.WriteLine("Обрабатывается книга: " + books.Key);
                var wb1 = new XLWorkbook(@"C:\Temp\" + books.Key + ".xlsx");
                var wb2 = new XLWorkbook(@"C:\Temp\" + books.Value + ".xlsx");

                for (int i = 11; i <= 179; i++)//18953
                {
                    for (int j = 11; j <= 179; j++)//18953
                    {
                        if (Convert.ToString(wb1.Worksheet(1).Row(i).Cell(1).Value).Trim() ==
                            Convert.ToString(wb2.Worksheet(1).Row(j).Cell(1).Value).Trim())
                        {
                            for (int k = 3; k <= 78; k++)
                            {
                                if (k <= 66 && k >= 59)
                                {
                                    Decimal d1 = Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(k).Value);
                                    Decimal d2 = Convert.ToDecimal(wb2.Worksheet(1).Row(j).Cell(k).Value);
                                    Decimal res = d1 + d2;
                                    wb1.Worksheet(1).Row(i).Cell(k).Value = res;
                                }

                            }
                            break;
                        }

                    }
                }
                wb1.Save();
                wb2.Save();
                Console.WriteLine("Сохранена книга: " + books.Key);
            }
        }

        public void DevideByTwo()
        {
            var wb1 = new XLWorkbook(@"C:\Temp\Согласие_итог_плюсом.xlsx");

            for (int i = 11; i <= 179; i++)//18953
            {

                for (int k = 3; k <= 52; k++)
                {
                    if (k == 4 || k == 6 || k == 10 || k == 12 || k == 16 || k == 18 || k == 24 || k == 26 || k == 30 || k == 32 || k == 36 || k == 38 || k == 44 || k == 46 || k == 50 || k == 52)
                    {
                        Decimal d1 = Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(k).Value);
                        Decimal res = d1 / 2;
                        wb1.Worksheet(1).Row(i).Cell(k).Value = res;
                    }

                }
            }
            wb1.Save();
        }
    }
}
