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
    class EzhkhInsertData
    {
        private pg pg;
        public EzhkhInsertData()
        {
            pg = new pg();
        }

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

        public void UpdateGkhCode()
        {
            Dictionary<int, string> районыСловарь = new Dictionary<int, string>();
            районыСловарь.Add(21654, "52");
            районыСловарь.Add(21655, "53");
            районыСловарь.Add(21656, "54");
            районыСловарь.Add(21657, "55");
            районыСловарь.Add(21658, "56");
            районыСловарь.Add(21659, "57");
            районыСловарь.Add(21660, "58");
            районыСловарь.Add(21661, "59");
            районыСловарь.Add(21662, "60");
            районыСловарь.Add(21663, "61");
            районыСловарь.Add(21664, "62");
            районыСловарь.Add(21665, "63");
            районыСловарь.Add(21666, "64");
            районыСловарь.Add(21667, "65");
            районыСловарь.Add(21668, "66");
            районыСловарь.Add(21669, "67");
            районыСловарь.Add(21670, "68");
            районыСловарь.Add(21671, "69");
            районыСловарь.Add(21672, "70");
            районыСловарь.Add(21673, "71");
            районыСловарь.Add(21674, "72");
            районыСловарь.Add(21675, "73");
            районыСловарь.Add(21676, "74");
            районыСловарь.Add(21677, "75");
            районыСловарь.Add(21678, "76");
            районыСловарь.Add(21679, "77");
            районыСловарь.Add(21680, "78");
            районыСловарь.Add(21682, "80");
            районыСловарь.Add(21683, "81");
            районыСловарь.Add(21684, "82");
            районыСловарь.Add(21685, "83");
            районыСловарь.Add(21686, "84");
            районыСловарь.Add(21687, "85");
            районыСловарь.Add(21688, "86");
            районыСловарь.Add(21689, "87");
            районыСловарь.Add(21690, "88");
            районыСловарь.Add(21691, "89");
            районыСловарь.Add(21692, "90");
            районыСловарь.Add(21693, "91");
            районыСловарь.Add(21694, "92");
            районыСловарь.Add(21695, "93");
            районыСловарь.Add(21696, "94");
            районыСловарь.Add(21697, "94");
            районыСловарь.Add(21698, "96");
            районыСловарь.Add(21699, "97");
            районыСловарь.Add(21700, "98");
            районыСловарь.Add(21701, "99");
            районыСловарь.Add(21702, "93");
            районыСловарь.Add(21681, "97");

            DataTable dt = pg.GetAllGkhCode();
            //int tempCode = 900;
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string gkhCode = dt.Rows[i][0].ToString();

                DataTable houseForUpdate = pg.GetHousesByGkhCode(gkhCode);
                for (int j = 0; j < houseForUpdate.Rows.Count; j++)
                {
                    if (gkhCode != "" && j == 0)
                        continue;
                    int id = Convert.ToInt32(houseForUpdate.Rows[j][0].ToString());
                    int minicipalityId = Convert.ToInt32(houseForUpdate.Rows[j][1].ToString());
                    if (minicipalityId == 21703)
                        continue;
                    int newGkhCode = pg.GetMaxGkhCodeByMunId(minicipalityId, районыСловарь[minicipalityId]);
                    pg.UpdateGkhCode(id, newGkhCode);
                    //tempCode++;
                }
            }
        }
    }
}
