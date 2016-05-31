using ClosedXML.Excel;
using ConsoleApplication1.Database;
using ConsoleApplication9;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1.mainCode
{
    class EzhkhInsertData
    {
        public void AddHouseManOrg()
        {
            EzhkhInsertDataDb ezhkhInsertDataDb = new EzhkhInsertDataDb();
            var wb2 = new XLWorkbook(@"C:\temp\МУП УД.xlsx");
            for (int i = 4; i <= 108; i++)
            {
                string str = ezhkhInsertDataDb.InsertHouseManOrg(
                    Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim(),
                    "02.03.2015", "17729406");
                if (str != "ЗАГРУЖЕНО")
                    Console.WriteLine(str);
            }
            wb2.Save();
        }

        public void AddHouseManOrgFromFile()
        {
            EzhkhInsertDataDb ezhkhInsertDataDb = new EzhkhInsertDataDb();
            var wb2 = new XLWorkbook(@"C:\temp\forImport.xlsx");
            for (int i = 3; i <= 9; i++)
            {
                if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value).Trim() != "")
                {
                    string str = ezhkhInsertDataDb.InsertHouseManOrg(
                    Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value).Trim(),
                    Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim().Substring(0, 10),
                    "17731219");
                    if (str != "ЗАГРУЖЕНО")
                        Console.WriteLine(str + "||1||" + i.ToString());
                }
            }
            for (int i = 4; i <= 6; i++)
            {
                if (Convert.ToString(wb2.Worksheet(2).Row(i).Cell(5).Value).Trim() != "")
                {
                    string str = ezhkhInsertDataDb.InsertHouseManOrg(
                    Convert.ToString(wb2.Worksheet(2).Row(i).Cell(5).Value).Trim(),
                    Convert.ToString(wb2.Worksheet(2).Row(i).Cell(4).Value).Trim().Substring(0, 10),
                    "17731219");
                    if (str != "ЗАГРУЖЕНО")
                        Console.WriteLine(str + "||2||" + i.ToString());
                }
            }
            for (int i = 4; i <= 30; i++)
            {
                if (Convert.ToString(wb2.Worksheet(3).Row(i).Cell(5).Value).Trim() != "")
                {
                    string str = ezhkhInsertDataDb.InsertHouseManOrg(
                    Convert.ToString(wb2.Worksheet(3).Row(i).Cell(5).Value).Trim(),
                    Convert.ToString(wb2.Worksheet(3).Row(i).Cell(4).Value).Trim().Substring(0, 10),
                    "17731219");
                    if (str != "ЗАГРУЖЕНО")
                        Console.WriteLine(str + "||3||" + i.ToString());
                }
            }
            for (int i = 3; i <= 68; i++)
            {
                if (Convert.ToString(wb2.Worksheet(4).Row(i).Cell(5).Value).Trim() != "")
                {
                    string str = ezhkhInsertDataDb.InsertHouseManOrg(
                    Convert.ToString(wb2.Worksheet(4).Row(i).Cell(5).Value).Trim(),
                    Convert.ToString(wb2.Worksheet(4).Row(i).Cell(4).Value).Trim().Substring(0, 10),
                    "17731219");
                    if (str != "ЗАГРУЖЕНО")
                        Console.WriteLine(str + "||4||" + i.ToString());
                }
            }
            wb2.Save();
        }

        public void AddCommunalOrg()
        {
            EzhkhInsertDataDb ezhkhInsertDataDb = new EzhkhInsertDataDb();
            string str = ezhkhInsertDataDb.InsertCommunalOrg("9548129");
            if (str != "ЗАГРУЖЕНО")
                Console.WriteLine(str);
        }

        public void AddResOrg()
        {
            EzhkhInsertDataDb ezhkhInsertDataDb = new EzhkhInsertDataDb();
            string str = ezhkhInsertDataDb.InsertResOrg("18512502");
            if (str != "ЗАГРУЖЕНО")
                Console.WriteLine(str);
        }

        public void UpdateAreaMkd()
        {
            EzhkhInsertDataDb ezhkhInsertDataDb = new EzhkhInsertDataDb();
            GetRoIdentifier getRoIdentifier = new GetRoIdentifier();
            var wb2 = new XLWorkbook(@"C:\temp\Копия Книга2.xlsx");
            for (int i = 2; i <= 6; i++)
            {
                if (i % 10 == 0)
                    Console.WriteLine(i);
                int roId = getRoIdentifier.SelectRoId("21702", wb2.Worksheet(1).Row(i).Cell(3).Value.ToString().Trim());
                if(roId == 0)
                {
                    wb2.Worksheet(1).Row(i).Cell(1).Style.Fill.BackgroundColor = XLColor.Yellow;
                    wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Yellow;
                    wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Yellow;
                    wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Yellow;
                }
                else
                {
                    wb2.Worksheet(1).Row(i).Cell(5).Value = roId;
                    string str = ezhkhInsertDataDb.UpdateMkdArea(roId, Convert.ToDecimal(wb2.Worksheet(1).Row(i).Cell(4).Value.ToString().Trim()));
                    if (str != "ЗАГРУЖЕНО")
                    {
                        wb2.Worksheet(1).Row(i).Cell(1).Style.Fill.BackgroundColor = XLColor.Red;
                        wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Red;
                        wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Red;
                        wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Red;
                    }
                }
               
            }
            wb2.Save();
        }
    }
}
