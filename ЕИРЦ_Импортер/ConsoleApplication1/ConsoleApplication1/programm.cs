using BytesRoad.Net.Ftp;
using ClosedXML.Excel;
using ConsoleApplication9;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Xml;

namespace ConsoleApplication1
{
    class Program
    {

        public static string Convert2(string value, Encoding src, Encoding trg)
        {
            Decoder dec = src.GetDecoder();
            byte[] ba = trg.GetBytes(value);
            int len = dec.GetCharCount(ba, 0, ba.Length);
            char[] ca = new char[len];
            dec.GetChars(ba, 0, ba.Length, ca, 0);
            return new string(ca);
        }

        public static Dictionary<string, int> SortMyDictionaryByKey(Dictionary<string, int> myDictionary)
        {
            List<KeyValuePair<string, int>> tempList = new List<KeyValuePair<string, int>>(myDictionary);
            tempList.Sort
            (
                delegate(KeyValuePair<string, int> firstPair, KeyValuePair<string, int> secondPair)
                {
                    return firstPair.Value.CompareTo(secondPair.Value);
                }
            );

            Dictionary<string, int> mySortedDictionary = new Dictionary<string, int>();
            foreach (KeyValuePair<string, int> pair in tempList)
            {
                mySortedDictionary.Add(pair.Key, pair.Value);
            }
            return mySortedDictionary;

        }

        public static Dictionary<string, decimal> SortMyDictionaryByKey(Dictionary<string, decimal> myDictionary)
        {
            List<KeyValuePair<string, decimal>> tempList = new List<KeyValuePair<string, decimal>>(myDictionary);
            tempList.Sort
            (
                delegate(KeyValuePair<string, decimal> firstPair, KeyValuePair<string, decimal> secondPair)
                {
                    return firstPair.Key.CompareTo(secondPair.Key);
                }
            );

            Dictionary<string, decimal> mySortedDictionary = new Dictionary<string, decimal>();
            foreach (KeyValuePair<string, decimal> pair in tempList)
            {
                mySortedDictionary.Add(pair.Key, pair.Value);
            }
            return mySortedDictionary;

        }

        static void Main(string[] args)
        {
            Ora ora = new Ora();
            Dbf dbf = new Dbf();
            pg pg = new pg();
            int type;
            Console.WriteLine("0 = Запись из эксельки: на вход только exhkh_code");
            Console.WriteLine("1 = Запись из эксельки:поставщик, житель, дом на основе ezhkh_code");
            Console.WriteLine("2 = Создание списка домов house.xlsx из Сервера");
            Console.WriteLine("3 = Билинг. Сверка домов из билинга с СЕРВЕРОМ");
            Console.WriteLine("4 = Билинг. Список нормативов");
            Console.WriteLine("5 = Билинг. Сопоставление кодов");
            Console.WriteLine("6 = Билинг. Количество домов по коду");
            Console.WriteLine("7 = Билинг");
            Console.WriteLine("8 = Билинг. Проставление название улицы");
            Console.WriteLine("9 = Билинг. Проверка штрих кода");
            Console.WriteLine("10 = Билинг. Дома из dbf");
            Console.WriteLine("11 = Формирование INSERT");
            Console.WriteLine("12 = Биллинг. Перерасчеты");
            Console.WriteLine("13 = Биллинг. Перерасчеты_2");
            Console.WriteLine("14 = Биллинг. Перерасчеты_3. Excel с кодами");
            Console.WriteLine("27 = Проставить ezhkh_code в Excel на основании другого файла");
            Console.WriteLine("41 = ЭЖКХ. Формирование Excel файла по текущему ремонту");
            Console.WriteLine("54 = ЭЖКХ. Перенос собственников по одному дому со SPLIT-ом по квартире");
            Console.WriteLine("44 = ЭЖКХ. Отчет по проценту заполнения домов");
            Console.WriteLine("55 = ЭЖКХ. Поставщики коммунальных услуг с добавление дома");
            Console.WriteLine("60 = Биллинг. Групповой ввод харакетритсик жилья");
            Console.WriteLine("68 = Биллинг. Перекидка");
            Console.WriteLine("71 = Стройка. Загрузка Договоров");
            Console.Write("Введите тип операции:");
            type = Convert.ToInt32(Console.ReadLine());

            #region 5
            if (type == 5)
            {
                string name = "normativ12";
                string name2 = "normativ_1";
                int rows = 1026;
                int rows2 = 1026;
                var wb = new XLWorkbook(@"C:\temp\" + name + ".xlsx");
                var wb2 = new XLWorkbook(@"C:\temp\" + name2 + ".xlsx");
                for (int i = 1; i <= rows; i++)
                {
                    for (int j = 2; j <= rows2; j++)
                    {
                        if (Convert.ToString(wb.Worksheet(1).Row(i).Cell(7).Value).Trim() == Convert.ToString(wb2.Worksheet(1).Row(j).Cell(1).Value).Trim())
                            //&& Convert.ToString(wb.Worksheet(1).Row(i).Cell(4).Value).Trim() == Convert.ToString(wb2.Worksheet(1).Row(j).Cell(3).Value).Trim())
                        {
                            wb.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                            //wb.Worksheet(1).Row(i).Cell(7).Value = Convert.ToString(wb2.Worksheet(1).Row(j).Cell(2).Value).Trim();
                            break;
                        }
                    }
                }
                wb.Save();
                wb2.Save();

            }
            #endregion

            #region 6
            else if (type == 6)
            {
                Dictionary<string, string> bases = new Dictionary<string, string>();
                bases.Add("BAZ_ALT", "78");
                bases.Add("BAZ_AYK", "79");
                bases.Add("BAZ_TREST", "82");
                bases.Add("baz_gks", "131");
                bases.Add("baz_rem", "130");
                Dictionary<string, int> kHouse = new Dictionary<string, int>();
                var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("1");
                foreach (KeyValuePair<string, string> db in bases)
                {
                    //int rows = 1;
                    string name = db.Key;
                    string dbPath = @"C:\imp\" + name + @"\spr\";
                    DataTable dtHouseTemp = dbf.SelectHouse(dbPath);
                    for (int i = 0; i < dtHouseTemp.Rows.Count; i++)
                    {
                        if (!kHouse.ContainsKey(dtHouseTemp.Rows[i][0].ToString()))
                        {
                            kHouse.Add(dtHouseTemp.Rows[i][0].ToString(), 1);
                        }
                        else
                        {
                            kHouse[dtHouseTemp.Rows[i][0].ToString()]++;
                        }
                    }
                }
                int row = 2;
                kHouse = SortMyDictionaryByKey(kHouse);
                foreach (KeyValuePair<string, int> val in kHouse)
                {
                    ws.Cell(row, 1).Value = val.Key;
                    ws.Cell(row, 2).Value = val.Value;
                    row++;
                }
                wb.SaveAs(@"C:\temp\bilHouseCOunt.xlsx");
            }
            #endregion

            #region 600
            else if (type == 600)
            {
                Dictionary<string, string> bases = new Dictionary<string, string>();
                bases.Add("BAZ_ALT", "78");
                Dictionary<string, int> kHouse = new Dictionary<string, int>();
                var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("1");
                int row = 2;
                foreach (KeyValuePair<string, string> db in bases)
                {
                    //int rows = 1;
                    string name = db.Key;
                    string dbPath = @"G:\work\base\Новая папка\Кировский\DEZ_VAS\g33";
                    DataTable dtHouseTemp = dbf.SelectKart(dbPath);
                    for (int i = 0; i < dtHouseTemp.Rows.Count; i++)
                    {
                        if (dtHouseTemp.Rows[i][3].ToString().Trim() == "517")
                        {
                            for (int j = 0; j < dtHouseTemp.Columns.Count; j++)
                            {
                                ws.Cell(row, j+1).Value = dtHouseTemp.Rows[i][j].ToString();
                            }
                            row++;
                        }
                    }
                }
                wb.SaveAs(@"C:\temp\bilHouseCOunt2.xlsx");
            }
            #endregion

            #region 7
            else if (type == 7)
            {
                Dictionary<string, string> bases = new Dictionary<string, string>();
                bases.Add("BAZ_ALT", "78");
                bases.Add("BAZ_AYK", "79");
                bases.Add("BAZ_TREST", "82");
                bases.Add("baz_gks", "131");
                bases.Add("baz_rem", "130");
                var wb2 = new XLWorkbook(@"C:\temp\bilHouseCOunt.xlsx");
                List<string> dt = new List<string>();
                for (int i = 2; i <= 102; i++)
                {
                    dt.Add(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim());
                }
                wb2.Save();
                List<string> kHouse = new List<string>();
                List<string> vu = new List<string>() { "0201", "0202" };
                var wb = new XLWorkbook();
                foreach (KeyValuePair<string, string> db in bases)
                {
                    int rows = 1;
                    string name = db.Key;
                    var ws = wb.Worksheets.Add(name);
                    string dbPath = @"C:\imp\" + name + @"\spr\";
                    string kodYK = "1111";
                    DataTable dtNormTemp = dbf.SelectNorm(dbPath);
                    DataTable dtHouseTemp = dbf.SelectHouse2(dbPath);
                    DataTable dtTarifTemp = dbf.SelectTarif(dbPath);
                    DataTable dtNorm = new DataTable();
                    DataTable dtHouse = new DataTable();
                    DataTable dtTarif = new DataTable();
                    dtNorm.Columns.Add("1");
                    dtNorm.Columns.Add("2");
                    dtNorm.Columns.Add("3");
                    dtNorm.Columns.Add("4");
                    dtHouse.Columns.Add("1");
                    dtHouse.Columns.Add("2");
                    dtHouse.Columns.Add("3");
                    dtHouse.Columns.Add("4");
                    dtHouse.Columns.Add("5");
                    dtTarif.Columns.Add("1");
                    dtTarif.Columns.Add("2");
                    dtTarif.Columns.Add("3");
                    DataRow row2;
                    DataRow row3;
                    DataRow row4;
                    DataTable dtHouseAll = dbf.SelectHouse(dbPath);
                    for (int i = 0; i < dtNormTemp.Rows.Count; i++)
                    {
                        if (vu.Contains(dtNormTemp.Rows[i][0].ToString()))
                        {
                            row2 = dtNorm.NewRow();
                            row2["1"] = dtNormTemp.Rows[i][0].ToString();
                            row2["2"] = dtNormTemp.Rows[i][1].ToString();
                            row2["3"] = dtNormTemp.Rows[i][2].ToString();
                            row2["4"] = dtNormTemp.Rows[i][3].ToString();
                            dtNorm.Rows.Add(row2);
                        }
                    }
                    for (int i = 0; i < dtHouseTemp.Rows.Count; i++)
                    {
                        if (vu.Contains(dtHouseTemp.Rows[i][1].ToString()))
                        {
                            row3 = dtHouse.NewRow();
                            if (dtHouseTemp.Rows[i][0].ToString() == "480")
                            {
                                
                            }
                            row3["1"] = dtHouseTemp.Rows[i][0].ToString();
                            row3["2"] = dtHouseTemp.Rows[i][1].ToString();
                            row3["3"] = dtHouseTemp.Rows[i][2].ToString();
                            row3["4"] = dtHouseTemp.Rows[i][3].ToString();
                            row3["5"] = dtHouseTemp.Rows[i][4].ToString();
                            dtHouse.Rows.Add(row3);
                        }
                    }
                    for (int i = 0; i < dtTarifTemp.Rows.Count; i++)
                    {
                        if (vu.Contains(dtTarifTemp.Rows[i][0].ToString()))
                        {
                            row4 = dtTarif.NewRow();
                            row4["1"] = dtTarifTemp.Rows[i][0].ToString();
                            row4["2"] = dtTarifTemp.Rows[i][1].ToString();
                            row4["3"] = dtTarifTemp.Rows[i][2].ToString();
                            dtTarif.Rows.Add(row4);
                        }
                    }
                    if (name == "baz_rem")
                    {
                        string str = "222";
                        str = str.Substring(1);
                    }
                    for (int i = 0; i < dtHouse.Rows.Count; i++)
                    {
                        if (!kHouse.Contains(dtHouse.Rows[i][0].ToString() + "|" + name))
                        {
                            if (dtHouse.Rows[i][0].ToString() == "931")
                            {
                                string sre = "etertertre";
                                sre.Substring(4);
                            }
                            if (dtHouse.Rows[i][2] != null && dtHouse.Rows[i][2].ToString() != "")
                            {
                                for (int j = 0; j < dtTarif.Rows.Count; j++)
                                {
                                    if (dtHouse.Rows[i][1].ToString() == dtTarif.Rows[j][0].ToString() && dtHouse.Rows[i][2].ToString() == dtTarif.Rows[j][1].ToString())
                                    {
                                        for (int k = 0; k < dtNorm.Rows.Count; k++)
                                        {
                                            if (dtNorm.Rows[k][0].ToString() == dtTarif.Rows[j][0].ToString() && dtNorm.Rows[k][2].ToString() == dtTarif.Rows[j][2].ToString())
                                            {
                                                ws.Cell(rows, 1).Value = dtHouse.Rows[i][0].ToString();
                                                string str = "";
                                                str = Convert2(dtNorm.Rows[k][1].ToString(), Encoding.GetEncoding(1251), Encoding.Default);
                                                ws.Cell(rows, 2).Value = str;
                                                //ws.Cell(i + 1, 2).Value = dtNormTemp.Rows[i][1].ToString();
                                                ws.Cell(rows, 3).Value = dtNorm.Rows[k][3].ToString();
                                                kHouse.Add(dtHouse.Rows[i][0].ToString() + "|" + name);
                                               
                                                for (int r = 0; r < dtHouseAll.Rows.Count; r++)
                                                {
                                                    if (dtHouseAll.Rows[r][0].ToString() == dtHouse.Rows[i][0].ToString())
                                                    {
                                                        ws.Cell(rows, 4).Value = dtHouseAll.Rows[r][1].ToString();
                                                        ws.Cell(rows, 5).Value = dtHouseAll.Rows[r][2].ToString();
                                                        ws.Cell(rows, 6).Value = dtHouseAll.Rows[r][3].ToString();
                                                        break;
                                                    }
                                                }
                                                rows++;
                                                break;
                                            }
                                        }
                                        break;
                                    }
                                }
                            }
                            else if (dtHouse.Rows[i][3] != null && dtHouse.Rows[i][3].ToString() != "")
                            {
                                for (int j = 0; j < dtTarif.Rows.Count; j++)
                                {
                                    if (dtHouse.Rows[i][1].ToString() == dtTarif.Rows[j][0].ToString() && dtHouse.Rows[i][3].ToString() == dtTarif.Rows[j][1].ToString())
                                    {
                                        for (int k = 0; k < dtNorm.Rows.Count; k++)
                                        {
                                            if (dtNorm.Rows[k][0].ToString() == dtTarif.Rows[j][0].ToString() && dtNorm.Rows[k][2].ToString() == dtTarif.Rows[j][2].ToString())
                                            {
                                                ws.Cell(rows, 1).Value = dtHouse.Rows[i][0].ToString();
                                                string str = "";
                                                str = Convert2(dtNorm.Rows[k][1].ToString(), Encoding.GetEncoding(1251), Encoding.Default);
                                                ws.Cell(rows, 2).Value = str;
                                                //ws.Cell(i + 1, 2).Value = dtNormTemp.Rows[i][1].ToString();
                                                ws.Cell(rows, 3).Value = dtNorm.Rows[k][3].ToString();
                                                kHouse.Add(dtHouse.Rows[i][0].ToString() + "|" + name);
                                                if (dtHouse.Rows[i][0].ToString() == "931")
                                                {
                                                    string sre = "etertertre";
                                                    sre.Substring(4);
                                                }
                                                
                                                for (int r = 0; r < dtHouseAll.Rows.Count; r++)
                                                {
                                                    if (dtHouseAll.Rows[r][0].ToString() == dtHouse.Rows[i][0].ToString())
                                                    {
                                                        ws.Cell(rows, 4).Value = dtHouseAll.Rows[r][1].ToString();
                                                        ws.Cell(rows, 5).Value = dtHouseAll.Rows[r][2].ToString();
                                                        ws.Cell(rows, 6).Value = dtHouseAll.Rows[r][3].ToString();
                                                        break;
                                                    }
                                                }
                                                rows++;
                                                break;
                                            }
                                        }
                                        break;
                                    }
                                }
                            }
                            else if (dtHouse.Rows[i][4] != null && dtHouse.Rows[i][4].ToString() != "")
                            {
                                for (int j = 0; j < dtTarif.Rows.Count; j++)
                                {
                                    if (dtHouse.Rows[i][1].ToString() == dtTarif.Rows[j][0].ToString() && dtHouse.Rows[i][4].ToString() == dtTarif.Rows[j][1].ToString())
                                    {
                                        for (int k = 0; k < dtNorm.Rows.Count; k++)
                                        {
                                            if (dtNorm.Rows[k][0].ToString() == dtTarif.Rows[j][0].ToString() && dtNorm.Rows[k][2].ToString() == dtTarif.Rows[j][2].ToString())
                                            {
                                                ws.Cell(rows, 1).Value = dtHouse.Rows[i][0].ToString();
                                                string str = "";
                                                str = Convert2(dtNorm.Rows[k][1].ToString(), Encoding.GetEncoding(1251), Encoding.Default);
                                                ws.Cell(rows, 2).Value = str;
                                                //ws.Cell(i + 1, 2).Value = dtNormTemp.Rows[i][1].ToString();
                                                ws.Cell(rows, 3).Value = dtNorm.Rows[k][3].ToString();
                                                kHouse.Add(dtHouse.Rows[i][0].ToString() + "|" + name);
                                                if (dtHouse.Rows[i][0].ToString() == "931")
                                                {
                                                    string sre = "etertertre";
                                                    sre.Substring(4);
                                                }
                                                for (int r = 0; r < dtHouseAll.Rows.Count; r++)
                                                {
                                                    if (dtHouseAll.Rows[r][0].ToString() == dtHouse.Rows[i][0].ToString())
                                                    {
                                                        ws.Cell(rows, 4).Value = dtHouseAll.Rows[r][1].ToString();
                                                        ws.Cell(rows, 5).Value = dtHouseAll.Rows[r][2].ToString();
                                                        ws.Cell(rows, 6).Value = dtHouseAll.Rows[r][3].ToString();
                                                        break;
                                                    }
                                                }
                                                rows++;
                                                break;
                                            }
                                        }
                                        break;
                                    }
                                }
                            }
                            else
                            {

                            }
                        }
                    }
                }
                wb.SaveAs(@"C:\temp\4\normativ.xlsx");
            }
            #endregion

            #region 8
            else if (type == 8)
            {
                var wb2 = new XLWorkbook(@"C:\temp\normativTemp.xlsx");
                List<string> dt = new List<string>();
                int rows = 1;
                string name = "BAZ_AYK";
                string dbPath = @"C:\imp\" + name + @"\spr\";
                DataTable dtStreet = dbf.SelectStreet(dbPath);
                for (int i = 2; i <= 150; i++)
                {
                    for(int j =1; j<dtStreet.Rows.Count;j++)
                    {
                        if (wb2.Worksheet(1).Row(i).Cell(3).Value.ToString() == dtStreet.Rows[j][3].ToString())
                        {
                            wb2.Worksheet(1).Row(i).Cell(3).Value = dtStreet.Rows[j][1].ToString();
                            break;
                        }
                    }
                }
                wb2.Save();
              
            }
            #endregion

            #region 9
            else if (type == 9)
            {
                string bar_code = "004019294045012061400004534931";

                int summ_kontr = Convert.ToInt32(bar_code.Substring(0, 1)) * 29 +
                Convert.ToInt32(bar_code.Substring(1, 1)) * 27 +
                Convert.ToInt32(bar_code.Substring(2, 1)) * 25 +
                Convert.ToInt32(bar_code.Substring(3, 1)) * 23 +
                Convert.ToInt32(bar_code.Substring(4, 1)) * 21 +
                Convert.ToInt32(bar_code.Substring(5, 1)) * 19 +
                Convert.ToInt32(bar_code.Substring(6, 1)) * 17 +
                Convert.ToInt32(bar_code.Substring(7, 1)) * 15 +
                Convert.ToInt32(bar_code.Substring(8, 1)) * 13 +
                Convert.ToInt32(bar_code.Substring(9, 1)) * 11 +
                Convert.ToInt32(bar_code.Substring(10, 1)) * 9 +
                Convert.ToInt32(bar_code.Substring(11, 1)) * 7 +
                Convert.ToInt32(bar_code.Substring(12, 1)) * 5 +
                Convert.ToInt32(bar_code.Substring(13, 1)) * 3 +
                Convert.ToInt32(bar_code.Substring(14, 1)) * 1 +
                Convert.ToInt32(bar_code.Substring(15, 1)) * 2 +
                Convert.ToInt32(bar_code.Substring(16, 1)) * 4 +
                Convert.ToInt32(bar_code.Substring(17, 1)) * 6 +
                Convert.ToInt32(bar_code.Substring(18, 1)) * 8 +
                Convert.ToInt32(bar_code.Substring(19, 1)) * 10 +
                Convert.ToInt32(bar_code.Substring(20, 1)) * 12 +
                Convert.ToInt32(bar_code.Substring(21, 1)) * 14 +
                Convert.ToInt32(bar_code.Substring(22, 1)) * 16 +
                Convert.ToInt32(bar_code.Substring(23, 1)) * 18 +
                Convert.ToInt32(bar_code.Substring(24, 1)) * 20 +
                Convert.ToInt32(bar_code.Substring(25, 1)) * 22 +
                Convert.ToInt32(bar_code.Substring(26, 1)) * 24 +
                Convert.ToInt32(bar_code.Substring(27, 1)) * 26;
                int rty = summ_kontr % 99;
                Console.WriteLine("2 цифры платежки = " + bar_code.Substring(28, 2));
                Console.WriteLine("контрольное число = " + rty);
            }
            #endregion

            #region 12- Перерасчеты
            else if (type == 12)
            {
                var wb2 = new XLWorkbook(@"C:\temp\2\unload_sam60.xlsx");
                DataRow row2;
                DataTable dt2 = new System.Data.DataTable();
                dt2.Columns.Add("1");
                dt2.Columns.Add("2");
                dt2.Columns.Add("3");
                Dictionary<string, string> convert = new Dictionary<string, string>();
                for (int i = 2; i <= 100000; i++)
                {
                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value) != "")
                    {
                        row2 = dt2.NewRow();
                        row2["1"] = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value);
                        row2["2"] = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value).PadLeft(5, '0');
                        row2["3"] = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value);
                        dt2.Rows.Add(row2);
                        convert.Add(row2["1"].ToString() + "|" + row2["2"].ToString() + "|" + Convert.ToString(wb2.Worksheet(1).Row(i).Cell(7).Value), row2["3"].ToString() + "|" + Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value));
                    }
                }
                Dictionary<string, string> files = new Dictionary<string, string>();
                /*files.Add("Альт-разовые-ХВС", "6");
                files.Add("Альт-разовые-водоотв", "7");
                files.Add("АУК-разовые-водоотв", "7");
                files.Add("ЖКС_ХВС_на_ГВС_Победы-8Г", "14");
                files.Add("ЖКС_ГВС_совАрмии23", "9");
                files.Add("ЖКС_ГВС_ЖЭУ35", "9");
                files.Add("ЖКС_ГВС_Балаков4_карякина2", "9");
                files.Add("ЖКС_ГВС_22Партс-10", "9");
                files.Add("ЖКС_ГВС_1Безымян-7", "9");
                files.Add("ЖКС_Водоотвед_совАрмии23", "7");
                files.Add("ЖКС_водоотвед_ЖЭУ35", "7");
                files.Add("ЖКС_Водоотвед_Балаков4_карякина2", "7");*/
                files.Add("Копия ЖКС-разовые-ГВС", "7");
                files.Add("Копия ЖКС-разовые-водоотв", "7");

                Dictionary<string, string> sup = new Dictionary<string, string>();
                sup.Add("ООО Самарские коммунальные системы", "612");
                sup.Add("ОАО ПТС", "974");
                sup.Add("ЗАО СамГЭС\"", "1039");
                sup.Add("ЗАО \"СамГЭС\"", "1039");
                sup.Add("ОАО \"Самараэнерго\"", "117");
                sup.Add("ООО \"Сбыт-Энерго\"", "289");
                sup.Add("ООО \"СВГК\"", "148");
                sup.Add("ОАО \"ВТГК\"", "100");
                sup.Add("ЗАО Коммунэнерго", "98");
                sup.Add("КЖКХ Советского р-на", "1042");
                sup.Add("ООО \"Жилищно-коммунальная система\"", "1071");
                string geu = "";
                StreamWriter sw = new StreamWriter(@"C:\temp\4\nedopost5.txt", false);
                foreach (KeyValuePair<string, string> fileName in files)
                {
                    sw.WriteLine(fileName.Key);
                    if (fileName.Key == "АУК")
                    {
                        string str = "55555";
                        str = str.Substring(5);
                    }
                    var wb = new XLWorkbook(@"C:\temp\3\" + fileName.Key + ".xlsx");
                    for (int i = 5; i <= 1000; i++)
                    {
                        if (wb.Worksheet(1).Row(i).Cell(1).Value.ToString() != "")
                        {
                            geu = wb.Worksheet(1).Row(i).Cell(1).Value.ToString();
                        }
                        if (wb.Worksheet(1).Row(i).Cell(3).Value.ToString() != "")
                        {
                            string strеееее = (Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(3).Value.ToString() + "|1";
                            string row = "";
                            //код ЛС из биллинга
                            if (convert.ContainsKey((Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(3).Value.ToString() + "|1"))
                            {
                                if (convert[(Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(3).Value.ToString() + "|1"].Split('|')[1].Contains(wb.Worksheet(1).Row(i).Cell(5).Value.ToString().Trim()))
                                {
                                    row += convert[(Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(3).Value.ToString() + "|1"].Split('|')[0] + "|";
                                }
                                else if (convert[(Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(3).Value.ToString() + "|0"].Split('|')[1].Contains(wb.Worksheet(1).Row(i).Cell(5).Value.ToString().Trim()))
                                {
                                    row += convert[(Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(3).Value.ToString() + "|0"].Split('|')[0] + "|";
                                }
                                else if (convert[(Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(3).Value.ToString() + "|2"].Split('|')[1].Contains(wb.Worksheet(1).Row(i).Cell(5).Value.ToString().Trim()))
                                {
                                    row += convert[(Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(3).Value.ToString() + "|2"].Split('|')[0] + "|";
                                }
                                else if (convert[(Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(3).Value.ToString() + "|3"].Split('|')[1].Contains(wb.Worksheet(1).Row(i).Cell(5).Value.ToString().Trim()))
                                {
                                    row += convert[(Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(3).Value.ToString() + "|3"].Split('|')[0] + "|";
                                }
                                else
                                {
                                    row += "!!!!!!~~~~|";
                                }
                            }
                            else
                            {
                                string str = (Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(3).Value.ToString() + "|0";
                                row += convert[(Convert.ToInt32(geu) + 800).ToString().Trim() + "|" + wb.Worksheet(1).Row(i).Cell(3).Value.ToString().Trim() + "|0"].Split('|')[0] + "|";
                            }
                            //код услуги
                            row += fileName.Value + "|";
                            //код поставщика
                            row += sup[wb.Worksheet(1).Row(3).Cell(1).Value.ToString()] + "|";
                            //месяц
                            row += "01.07.2014|";
                            //сумма перекидки
                            if (Convert.ToDecimal(wb.Worksheet(1).Row(i).Cell(6).Value) >= 0)
                            {
                                row += Math.Round(Convert.ToDecimal(wb.Worksheet(1).Row(i).Cell(6).Value), 2) * (-1) + "|";
                            }
                            else if (Convert.ToDecimal(wb.Worksheet(1).Row(i).Cell(10).Value) > 0)
                            {
                                row += Convert.ToDecimal(wb.Worksheet(1).Row(i).Cell(10).Value) * (-1) * (-1) + "|";
                            }
                            else
                            {
                                row += "0|";
                            }
                            //комментарий
                            row += "|";
                            //тариф
                            if (Convert.ToDecimal(wb.Worksheet(1).Row(i).Cell(6).Value) >= 0)
                            {
                                row += Math.Round(Convert.ToDecimal(wb.Worksheet(1).Row(i).Cell(6).Value), 2) * (-1) + "|";
                            }
                            else if (Convert.ToDecimal(wb.Worksheet(1).Row(i).Cell(10).Value) > 0)
                            {
                                row += Convert.ToDecimal(wb.Worksheet(1).Row(i).Cell(10).Value) * (-1) + "|";
                            }
                            else
                            {
                                row += "0|";
                            }
                            //расход
                            row += "1|";
                            sw.WriteLine(row);
                        }
                    }
                }
                sw.Close();
            }
            #endregion

            #region 13- Перерасчеты
            else if (type == 13)
            {
                var wb2 = new XLWorkbook(@"C:\temp\2\unload_sam60.xlsx");
                DataRow row2;
                DataTable dt2 = new System.Data.DataTable();
                dt2.Columns.Add("1");
                dt2.Columns.Add("2");
                dt2.Columns.Add("3");
                Dictionary<string, string> convert = new Dictionary<string, string>();
                for (int i = 2; i <= 100000; i++)
                {
                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value) != "")
                    {
                        row2 = dt2.NewRow();
                        row2["1"] = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value);
                        row2["2"] = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value).PadLeft(5, '0');
                        row2["3"] = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value);
                        dt2.Rows.Add(row2);
                        convert.Add(row2["1"].ToString() + "|" + row2["2"].ToString() + "|" + Convert.ToString(wb2.Worksheet(1).Row(i).Cell(7).Value), row2["3"].ToString() + "|" + Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value));
                    }
                }
                Dictionary<string, string> files = new Dictionary<string, string>();
                /*files.Add("Альт-разовые-ХВС", "6");
                files.Add("Альт-разовые-водоотв", "7");
                files.Add("АУК-разовые-водоотв", "7");
                files.Add("Копия ЖКС-разовые-водоотв", "7");
                files.Add("ЖКС_ГВС_совАрмии23", "9");
                files.Add("ЖКС_ГВС_ЖЭУ35", "9");
                files.Add("ЖКС_ГВС_Балаков4_карякина2", "9");
                files.Add("ЖКС_ГВС_22Партс-10", "9");
                files.Add("ЖКС_ГВС_1Безымян-7", "9");
                files.Add("ЖКС_Водоотвед_совАрмии23", "7");*/
                files.Add("Копия ЖКС-разовые-ХВС-на ГВС_компонент на хол.воду", "14");
                files.Add("Копия ЖКС-разовые-ХВС", "6");
                files.Add("Копия ЖКС-разовые-ГВС", "9");
                files.Add("Копия ЖКС-разовые-водоотв", "7");

                Dictionary<string, string> sup = new Dictionary<string, string>();
                sup.Add("ООО \"СКС\"", "612");
                sup.Add("ОАО \"ПТС\"", "974");
                sup.Add("ЗАО СамГЭС\"", "1039");
                sup.Add("ЗАО \"СамГЭС\"", "1039");
                sup.Add("ОАО \"Самараэнерго\"", "117");
                sup.Add("ООО \"Сбыт-Энерго\"", "289");
                sup.Add("ООО \"СВГК\"", "148");
                sup.Add("ОАО \"ВТГК\"", "100");
                sup.Add("ЗАО \"КОММУНЭНЕРГО\"", "98");
                sup.Add("КЖКХ Советского р-на", "1042");
                sup.Add("ООО \"Жилищно-коммунальная система\"", "1071");
                string geu = "";
                var wb3 = new XLWorkbook(@"C:\temp\5\Книга8.xlsx");
                DataRow row3;
                DataTable dt3 = new System.Data.DataTable();
                dt3.Columns.Add("1");
                dt3.Columns.Add("2");
                dt3.Columns.Add("3");
                dt3.Columns.Add("4");
                for (int i = 2; i <= 2499; i++)
                {
                    row3 = dt3.NewRow();
                    row3["1"] = Convert.ToString(wb3.Worksheet(1).Row(i).Cell(5).Value);
                    row3["2"] = Convert.ToString(wb3.Worksheet(1).Row(i).Cell(4).Value);
                    row3["3"] = Convert.ToString(wb3.Worksheet(1).Row(i).Cell(2).Value);
                    row3["4"] = Convert.ToString(wb3.Worksheet(1).Row(i).Cell(8).Value);
                    dt3.Rows.Add(row3);
                }
                foreach (KeyValuePair<string, string> fileName in files)
                {
                    var wb = new XLWorkbook(@"C:\temp\5\" + fileName.Key + ".xlsx");
                    for (int i = 8; i <= 1000; i++)
                    {
                        for (int j = 0; j < dt3.Rows.Count; j++)
                        {
                            if (wb.Worksheet(1).Row(i).Cell(7).Value.ToString() == dt3.Rows[j][0].ToString()
                                && wb.Worksheet(1).Row(i).Cell(6).Value.ToString() == dt3.Rows[j][1].ToString()
                                && wb.Worksheet(1).Row(i).Cell(5).Value.ToString() == dt3.Rows[j][2].ToString()
                                && (wb.Worksheet(1).Row(i).Cell(1).Value.ToString() == dt3.Rows[j][3].ToString()
                                    || (Convert.ToDecimal(wb.Worksheet(1).Row(i).Cell(2).Value) * (-1)).ToString() == dt3.Rows[j][3].ToString()))
                            {
                                wb.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                                break;
                            }
                        }
                        // проставляем коды в эксель
                        /*
                        if (wb.Worksheet(1).Row(i).Cell(3).Value.ToString() != "")
                        {
                            geu = wb.Worksheet(1).Row(i).Cell(3).Value.ToString();
                        }
                        if (wb.Worksheet(1).Row(i).Cell(3).Value.ToString() != "")
                        {
                            string strеееее = (Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(4).Value.ToString() + "|1";
                            string row = "";
                            //код ЛС из биллинга
                            if (convert.ContainsKey((Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(4).Value.ToString() + "|1"))
                            {
                                if (convert[(Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(4).Value.ToString() + "|1"].Split('|')[1].Contains(wb.Worksheet(1).Row(i).Cell(8).Value.ToString().Trim()))
                                {
                                    wb.Worksheet(1).Row(i).Cell(13).Value = convert[(Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(4).Value.ToString() + "|1"].Split('|')[0];
                                }
                                else if (convert[(Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(4).Value.ToString() + "|0"].Split('|')[1].Contains(wb.Worksheet(1).Row(i).Cell(8).Value.ToString().Trim()))
                                {
                                    wb.Worksheet(1).Row(i).Cell(13).Value = convert[(Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(4).Value.ToString() + "|0"].Split('|')[0];
                                }
                                else if (convert[(Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(4).Value.ToString() + "|2"].Split('|')[1].Contains(wb.Worksheet(1).Row(i).Cell(8).Value.ToString().Trim()))
                                {
                                    wb.Worksheet(1).Row(i).Cell(13).Value = convert[(Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(4).Value.ToString() + "|2"].Split('|')[0];
                                }
                                else if (convert[(Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(4).Value.ToString() + "|3"].Split('|')[1].Contains(wb.Worksheet(1).Row(i).Cell(8).Value.ToString().Trim()))
                                {
                                    wb.Worksheet(1).Row(i).Cell(13).Value = convert[(Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(4).Value.ToString() + "|3"].Split('|')[0];
                                }
                                else
                                {
                                    wb.Worksheet(1).Row(i).Cell(13).Value = "!!!!!!~~~~";
                                }
                            }
                            else
                            {
                                string str = (Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(4).Value.ToString() + "|0";
                                wb.Worksheet(1).Row(i).Cell(13).Value = convert[(Convert.ToInt32(geu) + 800).ToString().Trim() + "|" + wb.Worksheet(1).Row(i).Cell(4).Value.ToString().Trim() + "|0"].Split('|')[0];
                            }
                            wb.Worksheet(1).Row(i).Cell(14).Value = fileName.Value;
                            wb.Worksheet(1).Row(i).Cell(15).Value = sup[wb.Worksheet(1).Row(i).Cell(1).Value.ToString()];
                        }*/
                    }
                    wb.Save();
                }
            }
            #endregion

            #region 14- Перерасчеты
            else if (type == 14)
            {
                var wb2 = new XLWorkbook(@"C:\temp\2\unload_sam60.xlsx");
                DataRow row2;
                DataTable dt2 = new System.Data.DataTable();
                dt2.Columns.Add("1");
                dt2.Columns.Add("2");
                dt2.Columns.Add("3");
                Dictionary<string, string> convert = new Dictionary<string, string>();
                for (int i = 2; i <= 100000; i++)
                {
                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value) != "")
                    {
                        row2 = dt2.NewRow();
                        row2["1"] = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value);
                        row2["2"] = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value).PadLeft(5, '0');
                        row2["3"] = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value);
                        dt2.Rows.Add(row2);
                        convert.Add(row2["1"].ToString() + "|" + row2["2"].ToString() + "|" + Convert.ToString(wb2.Worksheet(1).Row(i).Cell(7).Value), row2["3"].ToString() + "|" + Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value));
                    }
                }
                Dictionary<string, string> files = new Dictionary<string, string>();
                /*files.Add("Альт-разовые-ХВС", "6");
                files.Add("Альт-разовые-водоотв", "7");
                files.Add("АУК-разовые-водоотв", "7");
                files.Add("ЖКС_ХВС_на_ГВС_Победы-8Г", "14");
                files.Add("ЖКС_ГВС_совАрмии23", "9");
                files.Add("ЖКС_ГВС_ЖЭУ35", "9");
                files.Add("ЖКС_ГВС_Балаков4_карякина2", "9");
                files.Add("ЖКС_ГВС_22Партс-10", "9");
                files.Add("ЖКС_ГВС_1Безымян-7", "9");
                files.Add("ЖКС_Водоотвед_совАрмии23", "7");
                files.Add("ЖКС_водоотвед_ЖЭУ35", "7");
                files.Add("ЖКС_Водоотвед_Балаков4_карякина2", "7");*/
                files.Add("Копия ЖКС-разовые-ХВС", "7");
                files.Add("Копия ЖКС-разовые-водоотв", "7");

                Dictionary<string, string> sup = new Dictionary<string, string>();
                sup.Add("ООО Самарские коммунальные системы", "612");
                sup.Add("ОАО ПТС", "974");
                sup.Add("ЗАО СамГЭС\"", "1039");
                sup.Add("ЗАО \"СамГЭС\"", "1039");
                sup.Add("ОАО \"Самараэнерго\"", "117");
                sup.Add("ООО \"Сбыт-Энерго\"", "289");
                sup.Add("ООО \"СВГК\"", "148");
                sup.Add("ОАО \"ВТГК\"", "100");
                sup.Add("ЗАО Коммунэнерго", "98");
                sup.Add("КЖКХ Советского р-на", "1042");
                sup.Add("ООО \"Жилищно-коммунальная система\"", "1071");
                string geu = "";
                StreamWriter sw = new StreamWriter(@"C:\temp\5\nedopost5_1Dop.txt", false);
                foreach (KeyValuePair<string, string> fileName in files)
                {
                    var wb = new XLWorkbook(@"C:\temp\5\" + fileName.Key + ".xlsx");
                    for (int i = 501; i <= 588; i++)
                    {
                        if (wb.Worksheet(1).Row(i).Cell(5).Value.ToString() != "")
                        {
                            string row = "";
                            //код ЛС из биллинга
                            row += wb.Worksheet(1).Row(i).Cell(5).Value.ToString().Trim() + "|";
                            //код услуги
                            row += wb.Worksheet(1).Row(i).Cell(6).Value.ToString().Trim() + "|";
                            //код поставщика
                            row += wb.Worksheet(1).Row(i).Cell(7).Value.ToString() + "|";
                            //месяц
                            row += "01.07.2014|";
                            //сумма перекидки
                            if (Convert.ToDecimal(wb.Worksheet(1).Row(i).Cell(1).Value) > 0)
                            {
                                row += Math.Round(Convert.ToDecimal(wb.Worksheet(1).Row(i).Cell(1).Value), 2) + "|";
                            }
                            else if (Convert.ToDecimal(wb.Worksheet(1).Row(i).Cell(2).Value) > 0)
                            {
                                row += Math.Round(Convert.ToDecimal(wb.Worksheet(1).Row(i).Cell(2).Value), 2) * (-1) + "|";
                            }
                            else
                            {
                                row += "0|";
                            }
                            //комментарий
                            row += wb.Worksheet(1).Row(i).Cell(4).Value.ToString() + "|";
                            sw.WriteLine(row);
                        }
                    }
                }
                sw.Close();
            }
            #endregion

            #region 11
            else if (type == 11)
            {
                /*insert into 
(no, nzp_conv_db, tabid, geu, kod, nzp, kodp, kodls_m, rc, nzp_was, is_open)
 values (1679563, 131, 3, '00', '969', 260681, '83', 0, 1, null, null);*/
                StreamWriter sw1 = new StreamWriter(@"C:\temp\5\kodPTS.txt", false);
                StreamWriter sw2 = new StreamWriter(@"C:\temp\5\kodVTGK.txt", false);
                StreamWriter sw3 = new StreamWriter(@"C:\temp\5\kodKomun.txt", false);
                StreamWriter sw4 = new StreamWriter(@"C:\temp\5\kodVolgo.txt", false);
                var wb1 = new XLWorkbook(@"C:\temp\5\Копия жксэ.xlsx");
                for (int i = 128; i <= 139; i++)
                {
                    if (wb1.Worksheet(1).Row(i).Cell(9).Value.ToString() == "ОАО ПТС")
                        sw1.Write(wb1.Worksheet(1).Row(i).Cell(4).Value.ToString() + ",");
                    else if (wb1.Worksheet(1).Row(i).Cell(9).Value.ToString() == "ООО Волгатеплоснаб")
                        sw4.Write(wb1.Worksheet(1).Row(i).Cell(4).Value.ToString() + ",");
                    else if (wb1.Worksheet(1).Row(i).Cell(9).Value.ToString() == "ЗАО Коммунэнерго")
                        sw3.Write(wb1.Worksheet(1).Row(i).Cell(4).Value.ToString() + ",");
                    else if (wb1.Worksheet(1).Row(i).Cell(9).Value.ToString() == "ОАО ВТГК")
                        sw2.Write(wb1.Worksheet(1).Row(i).Cell(4).Value.ToString() + ",");
                }
               
                sw1.Close();
                sw2.Close();
                sw3.Close();
            }
            #endregion

            #region 15
            else if (type == 15)
            {
                StreamWriter sw1 = new StreamWriter(@"C:\temp\5\ins2.txt", false);
                var wb1 = new XLWorkbook(@"C:\temp\5\mop.xlsx");
                for (int i = 2; i <= 65; i=i+2)
                {
                    if (wb1.Worksheet(1).Row(i).Cell(3).Value.ToString() == "11")
                    {
                        sw1.WriteLine("delete from  sam61_charge_14:charge_06 where nzp_charge = " + wb1.Worksheet(1).Row(i).Cell(1).Value.ToString() + " AND nzp_serv = 11;");
                        sw1.WriteLine("update sam61_charge_14:charge_06 SET sum_outsaldo = " +
                            (Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(4).Value.ToString()) + Convert.ToDecimal(wb1.Worksheet(1).Row(i + 1).Cell(4).Value.ToString())).ToString() + ", sum_pere = "+
                            (Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(5).Value.ToString()) + Convert.ToDecimal(wb1.Worksheet(1).Row(i + 1).Cell(5).Value.ToString())).ToString() + ", sum_insaldo = " +
                            (Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(6).Value.ToString()) + Convert.ToDecimal(wb1.Worksheet(1).Row(i + 1).Cell(6).Value.ToString())).ToString() + " where nzp_charge = " +
                            wb1.Worksheet(1).Row(i + 1).Cell(1).Value.ToString() + " AND nzp_serv = 17;");
                    }
                    else
                    {
                        sw1.WriteLine("delete from  sam61_charge_14:charge_06 where nzp_charge = " + wb1.Worksheet(1).Row(i+1).Cell(1).Value.ToString() + " AND nzp_serv = 11;");
                        sw1.WriteLine("update sam61_charge_14:charge_06 SET sum_outsaldo = " +
                            (Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(4).Value.ToString()) + Convert.ToDecimal(wb1.Worksheet(1).Row(i + 1).Cell(4).Value.ToString())).ToString() + ", sum_pere = " +
                            (Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(5).Value.ToString()) + Convert.ToDecimal(wb1.Worksheet(1).Row(i + 1).Cell(5).Value.ToString())).ToString() + ", sum_insaldo = " +
                            (Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(6).Value.ToString()) + Convert.ToDecimal(wb1.Worksheet(1).Row(i + 1).Cell(6).Value.ToString())).ToString() + " where nzp_charge = " +
                            wb1.Worksheet(1).Row(i).Cell(1).Value.ToString() + " AND nzp_serv = 17;");
                    }
                }

                sw1.Close();
            }
            #endregion

            #region 16- Перерасчеты
            else if (type == 16)
            {
                var wb2 = new XLWorkbook(@"C:\temp\2\unload_sam60.xlsx");
                DataRow row2;
                DataTable dt2 = new System.Data.DataTable();
                dt2.Columns.Add("1");
                dt2.Columns.Add("2");
                dt2.Columns.Add("3");
                Dictionary<string, string> convert = new Dictionary<string, string>();
                for (int i = 2; i <= 100000; i++)
                {
                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value) != "")
                    {
                        row2 = dt2.NewRow();
                        row2["1"] = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value);
                        row2["2"] = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value).PadLeft(5, '0');
                        row2["3"] = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value);
                        dt2.Rows.Add(row2);
                        convert.Add(row2["1"].ToString() + "|" + row2["2"].ToString() + "|" + Convert.ToString(wb2.Worksheet(1).Row(i).Cell(7).Value), row2["3"].ToString() + "|" + Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value));
                    }
                }
                Dictionary<string, string> files = new Dictionary<string, string>();
                /*files.Add("Альт-разовые-ХВС", "6");
                files.Add("Альт-разовые-водоотв", "7");
                files.Add("АУК-разовые-водоотв", "7");
                files.Add("ЖКС_ХВС_на_ГВС_Победы-8Г", "14");
                files.Add("ЖКС_ГВС_совАрмии23", "9");
                files.Add("ЖКС_ГВС_ЖЭУ35", "9");
                files.Add("ЖКС_ГВС_Балаков4_карякина2", "9");
                files.Add("ЖКС_ГВС_22Партс-10", "9");
                files.Add("ЖКС_ГВС_1Безымян-7", "9");
                files.Add("ЖКС_Водоотвед_совАрмии23", "7");
                files.Add("ЖКС_водоотвед_ЖЭУ35", "7");
                files.Add("ЖКС_Водоотвед_Балаков4_карякина2", "7");*/
                files.Add("_ОДН-ЖКС", "7");

                Dictionary<string, string> sup = new Dictionary<string, string>();
                sup.Add("ООО Самарские коммунальные системы", "612");
                sup.Add("ОАО ПТС", "974");
                sup.Add("ЗАО СамГЭС\"", "1039");
                sup.Add("ЗАО \"СамГЭС\"", "1039");
                sup.Add("ОАО \"Самараэнерго\"", "117");
                sup.Add("ООО \"Сбыт-Энерго\"", "289");
                sup.Add("ООО \"СВГК\"", "148");
                sup.Add("ОАО \"ВТГК\"", "100");
                sup.Add("ЗАО Коммунэнерго", "98");
                sup.Add("КЖКХ Советского р-на", "1042");
                sup.Add("ООО \"Жилищно-коммунальная система\"", "1071");
                string geu = "";
                foreach (KeyValuePair<string, string> fileName in files)
                {
                    var wb = new XLWorkbook(@"C:\temp\5\" + fileName.Key + ".xlsx");
                    for (int i = 4; i <= 49771; i++)
                    {
                        if (wb.Worksheet(1).Row(i).Cell(1).Value.ToString() != "")
                        {
                            geu = wb.Worksheet(1).Row(i).Cell(1).Value.ToString();
                        }
                        if (wb.Worksheet(1).Row(i).Cell(3).Value.ToString() != "")
                        {
                            try
                            {
                                string strеееее = (Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(3).Value.ToString().Trim() + "|1";
                                string row = "";
                                //код ЛС из биллинга
                                if (convert.ContainsKey((Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(3).Value.ToString().Split(' ')[0].Trim() + "|1"))
                                {
                                    if (convert[(Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(3).Value.ToString().Split(' ')[0].Trim() + "|1"].Split('|')[1].Contains(wb.Worksheet(1).Row(i).Cell(5).Value.ToString().Trim().Split(' ')[0].Trim()))
                                    {
                                        wb.Worksheet(1).Row(i).Cell(8).Value = convert[(Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(3).Value.ToString().Split(' ')[0].Trim() + "|1"].Split('|')[0];
                                    }
                                    else if (convert[(Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(3).Value.ToString().Split(' ')[0].Trim() + "|0"].Split('|')[1].Contains(wb.Worksheet(1).Row(i).Cell(5).Value.ToString().Trim().Split(' ')[0].Trim()))
                                    {
                                        wb.Worksheet(1).Row(i).Cell(8).Value = convert[(Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(3).Value.ToString().Split(' ')[0].Trim() + "|0"].Split('|')[0];
                                    }
                                    else if (convert[(Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(3).Value.ToString().Split(' ')[0].Trim() + "|2"].Split('|')[1].Contains(wb.Worksheet(1).Row(i).Cell(5).Value.ToString().Trim().Split(' ')[0].Trim()))
                                    {
                                        wb.Worksheet(1).Row(i).Cell(8).Value = convert[(Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(3).Value.ToString().Split(' ')[0].Trim() + "|2"].Split('|')[0];
                                    }
                                    else if (convert[(Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(3).Value.ToString().Split(' ')[0].Trim() + "|3"].Split('|')[1].Contains(wb.Worksheet(1).Row(i).Cell(5).Value.ToString().Trim().Split(' ')[0].Trim()))
                                    {
                                        wb.Worksheet(1).Row(i).Cell(8).Value = convert[(Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(3).Value.ToString().Split(' ')[0].Trim() + "|3"].Split('|')[0];
                                    }
                                    else
                                    {
                                        wb.Worksheet(1).Row(i).Cell(8).Value = "!!!!!!~~~~";
                                        wb.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                                    }
                                }
                                else
                                {
                                    string str = (Convert.ToInt32(geu) + 800).ToString() + "|" + wb.Worksheet(1).Row(i).Cell(3).Value.ToString().Split(' ')[0].Trim() + "|0";
                                    wb.Worksheet(1).Row(i).Cell(8).Value = convert[(Convert.ToInt32(geu) + 800).ToString().Trim() + "|" + wb.Worksheet(1).Row(i).Cell(3).Value.ToString().Trim().Split(' ')[0].Trim() + "|0"].Split('|')[0];
                                }
                            }
                            catch
                            {
                                wb.Worksheet(1).Row(i).Cell(8).Value = "````~~~~~~~~~~~";
                                wb.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                            }
                            
                        }
                    }
                    wb.Save();
                }
            }
            #endregion

            #region 17
            else if (type == 17)
            {
                StreamWriter sw1 = new StreamWriter(@"C:\temp\5\upload4.txt", false);
                var wb1 = new XLWorkbook(@"C:\temp\5\_ОДН-ЖКС.xlsx");
                for (int i = 5; i <= 60000; i++)
                {
                    if (wb1.Worksheet(1).Row(i).Cell(3).Value.ToString() != ""
                        && (wb1.Worksheet(1).Row(i).Cell(6).Value != null || wb1.Worksheet(1).Row(i).Cell(7).Value != null)
                        && ((wb1.Worksheet(1).Row(i).Cell(6).Value.ToString() != "" && wb1.Worksheet(1).Row(i).Cell(6).Value.ToString() != "0")
                        || (wb1.Worksheet(1).Row(i).Cell(7).Value.ToString() != "" && wb1.Worksheet(1).Row(i).Cell(7).Value.ToString() != "0")))
                    {
                        string row = "";
                        row += wb1.Worksheet(1).Row(i).Cell(8).Value.ToString() + "|";
                        if (wb1.Worksheet(1).Row(i).Cell(6).Value.ToString() != "" && wb1.Worksheet(1).Row(i).Cell(6).Value.ToString() != "0")
                            row += Math.Round(Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(6).Value.ToString()), 2) + "|";
                        else
                            row += "0|";
                        if (wb1.Worksheet(1).Row(i).Cell(7).Value.ToString() != "" && wb1.Worksheet(1).Row(i).Cell(7).Value.ToString() != "0")
                            row += Math.Round(Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(7).Value.ToString()), 2) + "|";
                        else
                            row += "0|";
                        sw1.WriteLine(row);
                    }
                }
                sw1.Close();
            }
            #endregion

            #region 20
            else if (type == 20)
            {
                StreamWriter sw1 = new StreamWriter(@"C:\temp\5\upload5.txt", false);
                var wb1 = new XLWorkbook(@"C:\temp\5\Копия ржу_сальдо ээ.xlsx");
                for (int i = 3; i <= 181; i++)
                {
                    if (wb1.Worksheet(1).Row(i).Cell(1).Value.ToString() != ""
                        && (wb1.Worksheet(1).Row(i).Cell(4).Value != null)
                        && (wb1.Worksheet(1).Row(i).Cell(4).Value.ToString() != "0"))
                    {
                        string row = "";
                        row += wb1.Worksheet(1).Row(i).Cell(6).Value.ToString() + "|";
                        if (wb1.Worksheet(1).Row(i).Cell(4).Value.ToString() != "" && wb1.Worksheet(1).Row(i).Cell(4).Value.ToString() != "0")
                            row += Math.Round(Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(4).Value.ToString()), 2) + "|0|";
                        sw1.WriteLine(row);
                    }
                }
                sw1.Close();
            }
            #endregion

            #region 18- Перерасчеты
            else if (type == 18)
            {
                
                Dictionary<string, string> files = new Dictionary<string, string>();
                /*files.Add("Альт-разовые-ХВС", "6");
                files.Add("Альт-разовые-водоотв", "7");
                files.Add("АУК-разовые-водоотв", "7");
                files.Add("ЖКС_ХВС_на_ГВС_Победы-8Г", "14");
                files.Add("ЖКС_ГВС_совАрмии23", "9");
                files.Add("ЖКС_ГВС_ЖЭУ35", "9");
                files.Add("ЖКС_ГВС_Балаков4_карякина2", "9");
                files.Add("ЖКС_ГВС_22Партс-10", "9");
                files.Add("ЖКС_ГВС_1Безымян-7", "9");
                files.Add("ЖКС_Водоотвед_совАрмии23", "7");*/
                files.Add("_ОДН-АУК", "7");
                files.Add("_ОДН-ПЖРТ", "7");
                files.Add("Копия _ОДН-Альт-2", "7");
                decimal sum1 = 0;
                decimal sum2 = 0;

                foreach (KeyValuePair<string, string> fileName in files)
                {
                    var wb = new XLWorkbook(@"C:\temp\5\" + fileName.Key + ".xlsx");
                    for (int i = 4; i <= 49771; i++)
                    {
                        if (wb.Worksheet(1).Row(i).Cell(3).Value.ToString() != "")
                        {
                            if (wb.Worksheet(1).Row(i).Cell(6).Value.ToString() != "" && wb.Worksheet(1).Row(i).Cell(6).Value.ToString() != "0")
                                sum1 += Math.Round(Convert.ToDecimal(wb.Worksheet(1).Row(i).Cell(6).Value.ToString()), 2);

                            if (wb.Worksheet(1).Row(i).Cell(7).Value.ToString() != "" && wb.Worksheet(1).Row(i).Cell(7).Value.ToString() != "0")
                                sum2 += Math.Round(Convert.ToDecimal(wb.Worksheet(1).Row(i).Cell(7).Value.ToString()), 2);
                        }
                    }
                    wb.Save();
                }
                Console.WriteLine(sum1);
                Console.WriteLine(sum2);
            }
            #endregion

            #region 19- Перерасчеты
            else if (type == 19)
            {
                var wb2 = new XLWorkbook(@"C:\temp\2\unload_sam60.xlsx");
                DataRow row2;
                DataTable dt2 = new System.Data.DataTable();
                dt2.Columns.Add("1");
                dt2.Columns.Add("2");
                dt2.Columns.Add("3");
                Dictionary<string, string> convert = new Dictionary<string, string>();
                for (int i = 2; i <= 100000; i++)
                {
                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value) != "" && Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value) == "892")
                    {
                        row2 = dt2.NewRow();
                        row2["1"] = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value);
                        row2["2"] = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value).PadLeft(5, '0');
                        row2["3"] = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value);
                        dt2.Rows.Add(row2);
                        convert.Add(row2["2"].ToString() + "|" + Convert.ToString(wb2.Worksheet(1).Row(i).Cell(7).Value), Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value));
                    }
                }
                Dictionary<string, string> files = new Dictionary<string, string>();
                /*files.Add("Альт-разовые-ХВС", "6");
                files.Add("Альт-разовые-водоотв", "7");
                files.Add("АУК-разовые-водоотв", "7");
                files.Add("ЖКС_ХВС_на_ГВС_Победы-8Г", "14");
                files.Add("ЖКС_ГВС_совАрмии23", "9");
                files.Add("ЖКС_ГВС_ЖЭУ35", "9");
                files.Add("ЖКС_ГВС_Балаков4_карякина2", "9");
                files.Add("ЖКС_ГВС_22Партс-10", "9");
                files.Add("ЖКС_ГВС_1Безымян-7", "9");
                files.Add("ЖКС_Водоотвед_совАрмии23", "7");
                files.Add("ЖКС_водоотвед_ЖЭУ35", "7");
                files.Add("ЖКС_Водоотвед_Балаков4_карякина2", "7");*/
                files.Add("Копия ржу_сальдо ээ", "7");

                Dictionary<string, string> sup = new Dictionary<string, string>();
                sup.Add("ООО Самарские коммунальные системы", "612");
                sup.Add("ОАО ПТС", "974");
                sup.Add("ЗАО СамГЭС\"", "1039");
                sup.Add("ЗАО \"СамГЭС\"", "1039");
                sup.Add("ОАО \"Самараэнерго\"", "117");
                sup.Add("ООО \"Сбыт-Энерго\"", "289");
                sup.Add("ООО \"СВГК\"", "148");
                sup.Add("ОАО \"ВТГК\"", "100");
                sup.Add("ЗАО Коммунэнерго", "98");
                sup.Add("КЖКХ Советского р-на", "1042");
                sup.Add("ООО \"Жилищно-коммунальная система\"", "1071");
                string geu = "";
                foreach (KeyValuePair<string, string> fileName in files)
                {
                    var wb = new XLWorkbook(@"C:\temp\5\" + fileName.Key + ".xlsx");
                    for (int i = 3; i <= 181; i++)
                    {
                        if (wb.Worksheet(1).Row(i).Cell(1).Value.ToString() != "")
                        {
                            try
                            {
                                //string strеееее = wb.Worksheet(1).Row(i).Cell(1).Value.ToString().Trim() + "|0";
                                string row = "";
                                //код ЛС из биллинга
                                if (convert.ContainsKey(wb.Worksheet(1).Row(i).Cell(1).Value.ToString().Split(' ')[0].Trim().PadLeft(5, '0') + "|1"))
                                {
                                        wb.Worksheet(1).Row(i).Cell(6).Value = "!!!!!!~~~~";
                                        wb.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                                }
                                else
                                {
                                    //string str = wb.Worksheet(1).Row(i).Cell(1).Value.ToString().Split(' ')[0].Trim() + "|0";
                                    wb.Worksheet(1).Row(i).Cell(6).Value = convert[wb.Worksheet(1).Row(i).Cell(1).Value.ToString().Trim().Split(' ')[0].Trim().PadLeft(5, '0') + "|0"];
                                }
                            }
                            catch
                            {
                                wb.Worksheet(1).Row(i).Cell(6).Value = "````~~~~~~~~~~~";
                                wb.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                            }

                        }
                    }
                    wb.Save();
                }
            }
            #endregion

            #region 10
            else if (type == 10)
            {
                Dictionary<string, string> bases = new Dictionary<string, string>();
                bases.Add("BAZ_ALT", "78");
                bases.Add("BAZ_AYK", "79");
                bases.Add("BAZ_TREST", "82");
                bases.Add("baz_gks", "131");
                bases.Add("baz_rem", "130");
                StreamWriter sw = new StreamWriter(@"C:\temp\1\Грузить.txt", false);
                StreamWriter sw2 = new StreamWriter(@"C:\temp\1\Грузить_ЖЭУ_УД_.txt", false);
                StreamWriter sw3 = new StreamWriter(@"C:\temp\1\Коды.txt", false);
                foreach (KeyValuePair<string, string> db in bases)
                {
                    string name = db.Key;
                    string dbPath = @"C:\imp\" + name + @"\spr\";
                    string kodYK = "1111";
                    DataTable dtKodIN = dbf.SelectKodIN(dbPath);
                    DataTable dtHouse = dbf.SelectHouse(dbPath, "1");
                    DataTable dtStreet = dbf.SelectStreet(dbPath, "1");
                    Dictionary<string, decimal> dtArea = dbf.SelectArea(dbPath, "1");
                    for (int i = 0; i < dtKodIN.Rows.Count; i++)
                    {
                        if (dtKodIN.Rows[i][2].ToString() == "0")
                        {
                            string rowWrite = "";
                            string rowWrite2 = "";
                            string rowWrite3 = "";

                            for (int j = 0; j < dtHouse.Rows.Count; j++)
                            {
                                if (dtHouse.Rows[j][6].ToString() == dtKodIN.Rows[i][0].ToString())
                                {
                                    for (int k = 0; k < dtStreet.Rows.Count; k++)
                                    {
                                        if (dtStreet.Rows[k][1].ToString() == dtHouse.Rows[j][2].ToString())
                                        {
                                            rowWrite2 += db.Value + "|";
                                            rowWrite2 += dtKodIN.Rows[i][0].ToString() + "|";
                                            rowWrite2 += dtHouse.Rows[j][2].ToString() + "|";
                                            //1
                                            rowWrite += "3||";
                                            //rowWrite2 += "3||";
                                            //3
                                            rowWrite += dtKodIN.Rows[i][0].ToString() + "|";
                                            //rowWrite2 += dtKodIN.Rows[i][0].ToString() + "|";
                                            //4
                                            rowWrite += "г.о. Самара||";
                                            //rowWrite2 += "г.о. Самара||";
                                            //6
                                            rowWrite += dtStreet.Rows[k][0].ToString() + "|";
                                            //rowWrite2 += dtStreet.Rows[k][0].ToString() + "|";
                                            //7
                                            rowWrite += dtHouse.Rows[j][0].ToString().Replace('└', 'А').Replace('┬', 'В') + "||" + kodYK + "|";
                                            //rowWrite2 += dtHouse.Rows[j][0].ToString() + "||" + kodYK + "|";
                                            //10
                                            rowWrite += dtKodIN.Rows[i][1].ToString() + "|";
                                            //rowWrite2 += dtKodIN.Rows[i][1].ToString() + "|";
                                            //11
                                            rowWrite += dtHouse.Rows[j][3].ToString() + "|";
                                            // rowWrite2 += dtHouse.Rows[j][3].ToString() + "|";
                                            //12
                                            if (dtHouse.Rows[j][1] == null || dtHouse.Rows[j][1].ToString() == "")
                                            {
                                                rowWrite += "01.01.1950" + "|";
                                                //rowWrite2 += "01.01.1950" + "|";
                                            }
                                            else
                                            {
                                                rowWrite += "01.01." + dtHouse.Rows[j][1].ToString() + "|";
                                                // rowWrite2 += "01.01." + dtHouse.Rows[j][1].ToString() + "|";
                                            }
                                            //13
                                            rowWrite += dtArea[dtKodIN.Rows[i][0].ToString()].ToString() + "|||96||||0|";
                                            //rowWrite2 += dtArea[dtKodIN.Rows[i][0].ToString()].ToString() + "|||96||||0|";
                                            //temp
                                            //rowWrite2 += dtHouse.Rows[j][4].ToString() + "|";
                                            //rowWrite2 += dtHouse.Rows[j][5].ToString() + "|";

                                            sw.WriteLine(rowWrite);
                                            sw2.WriteLine(rowWrite2);
                                            break;
                                        }
                                    }
                                    break;
                                }
                            }


                        }
                    }
                }
                sw.Close();
                sw2.Close();
            }
            #endregion

            #region 21
            else if (type == 21)
            {
                Dictionary<string, string> bases = new Dictionary<string, string>();
                bases.Add("BAZA", "78");
                bases.Add("Baza_GKS", "79");
                bases.Add("BAZA_PTS", "82");
                List<string> res = new List<string>();
                int countAll = 0;
                foreach (KeyValuePair<string, string> db in bases)
                {
                    int count = 0;
                    string name = db.Key;
                    string dbPath = @"C:\server_sam\" + name + @"\spr\";
                    DataTable dtHouse = dbf.SelectHouse(dbPath);
                    for (int i = 0; i < dtHouse.Rows.Count; i++)
                    {
                        count++;
                        countAll++;
                    }
                    res.Add(name + " =" + count);
                }
                res.Add("All" + " =" + countAll);
                foreach (string str in res)
                {
                    Console.WriteLine(str);
                }
            }
            #endregion

            #region 22
            else if (type == 22)
            {
                Dictionary<string, string> bases = new Dictionary<string, string>();
                
                bases.Add(@"Baza_GKS\g41", "Baza_GKS");
                bases.Add(@"Baza_GKS\g43", "Baza_GKS");
                bases.Add(@"Baza_GKS\g46", "Baza_GKS");
                bases.Add(@"BAZA_PTS\g01", "BAZA_PTS");
                bases.Add(@"BAZA_PTS\g03", "BAZA_PTS");
                bases.Add(@"BAZA_PTS\g06", "BAZA_PTS");
                bases.Add(@"BAZA\g11", "BAZA");
                bases.Add(@"BAZA\g33", "BAZA");
                bases.Add(@"BAZA\g66", "BAZA");
                List<string> res = new List<string>();
                int countAll = 0;
                decimal areaAll1 = 0;
                decimal areaAll2 = 0;
                int kolGAll1 = 0;
                int kolGAll2 = 0;
                string tmpName = "";
                int count = 0;
                decimal area1 = 0;
                decimal area2 = 0;
                int kolG1 = 0;
                int kolG2 = 0;
                bool first = true;
                foreach (KeyValuePair<string, string> db in bases)
                {

                    string name = db.Key;
                    string dbPath = @"C:\server_sam\" + name + @"\";
                    if (db.Value != tmpName)
                    {
                        if (!first)
                        {
                            res.Add(tmpName + ". ЛС = " + count);
                            res.Add(tmpName + ". ПЛошадь1 = " + area1);
                            res.Add(tmpName + ". ПЛошадь2 = " + area2);
                            res.Add(tmpName + ". Жильцы1 = " + kolG1);
                            res.Add(tmpName + ". Жильцы2 = " + kolG2);
                        }
                        first = false;
                        tmpName = db.Value;
                        count = 0;
                        area1 = 0;
                        area2 = 0;
                        kolG1 = 0;
                        kolG2 = 0;
                    }
                    DataTable dtHouse = dbf.SelectLS(dbPath);
                    for (int i = 0; i < dtHouse.Rows.Count; i++)
                    {
                        count++;
                        area1 += Convert.ToDecimal(dtHouse.Rows[i][1].ToString());
                        area2 += Convert.ToDecimal(dtHouse.Rows[i][2].ToString());
                        if (dtHouse.Rows[i][3] != null && dtHouse.Rows[i][3].ToString() != "")
                        {
                            kolG1 += Convert.ToInt32(dtHouse.Rows[i][3].ToString());
                            kolGAll1 += Convert.ToInt32(dtHouse.Rows[i][3].ToString());
                        }
                        if (dtHouse.Rows[i][4] != null && dtHouse.Rows[i][4].ToString() != "")
                        {
                            kolG2 += Convert.ToInt32(dtHouse.Rows[i][4].ToString());
                            kolGAll2 += Convert.ToInt32(dtHouse.Rows[i][4].ToString());
                        }

                        countAll++;
                        areaAll1 += Convert.ToDecimal(dtHouse.Rows[i][1].ToString());
                        areaAll2 += Convert.ToDecimal(dtHouse.Rows[i][2].ToString());
                    }
                    
                   
                }
                res.Add("BAZA. ЛС = " + count);
                res.Add("BAZA. ПЛошадь1 = " + area1);
                res.Add("BAZA. ПЛошадь2 = " + area2);
                res.Add("BAZA. Жильцы1 = " + kolG1);
                res.Add("BAZA. Жильцы2 = " + kolG2);
                res.Add("All . ЛС" + " =" + countAll);
                res.Add("All . ПЛошадь1" + " =" + areaAll1);
                res.Add("All . ПЛошадь2" + " =" + areaAll2);
                res.Add("All . Жильцы1" + " =" + kolGAll1);
                res.Add("All . Жильцы2" + " =" + kolGAll1);
                foreach (string str in res)
                {
                    Console.WriteLine(str);
                }
            }
            #endregion

            #region 23
            else if (type == 23)
            {
                Dictionary<string, string> bases = new Dictionary<string, string>();
                bases.Add(@"Baza\g11", "Baza");
                bases.Add(@"Baza\g33", "Baza");
                bases.Add(@"Baza\g66", "Baza");
                var wb = new XLWorkbook(@"C:\temp\6\Книга1.xlsx");
                DataTable dt = new DataTable();
                dt.Columns.Add("1");
                dt.Columns.Add("2");
                dt.Columns.Add("3");
                DataRow row2;
                Dictionary<string, string> convert = new Dictionary<string, string>();
                List<string> ls = new List<string>();
                for(int i= 2;i<=9530;i++)
                {
                    row2 = dt.NewRow();
                    row2["1"] = Convert.ToString(wb.Worksheet(1).Row(i).Cell(1).Value);
                    row2["2"] = Convert.ToString(wb.Worksheet(1).Row(i).Cell(2).Value);
                    row2["3"] = Convert.ToString(wb.Worksheet(1).Row(i).Cell(3).Value).PadLeft(5, '0');
                    dt.Rows.Add(row2);
                }
                Dictionary<string, decimal> res = new Dictionary<string, decimal>();
                foreach (KeyValuePair<string, string> db in bases)
                {

                    string name = db.Key;
                    string dbPath = @"C:\server_sam\" + name + @"\";
                    DataTable dtHouse = dbf.SelectKart(dbPath);
                    for (int i = 0; i < dtHouse.Rows.Count; i++)
                    {
                        bool isFind = false;
                        for (int j = 0; j < dt.Rows.Count; j++)
                        {
                            if (dt.Rows[j][0].ToString() == dtHouse.Rows[i][0].ToString() &&
                                dt.Rows[j][2].ToString() == dtHouse.Rows[i][4].ToString())
                            {
                                isFind = true;
                                break;
                            }
                            
                        }
                        if (!isFind)
                            ls.Add(dtHouse.Rows[i][0].ToString() + "|" + dtHouse.Rows[i][4].ToString());
                    }
                }

                foreach (string str in ls)
                {
                    Console.WriteLine(str);
                }

               /* var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("1");
                int row = 2;
                res = SortMyDictionaryByKey(res);
                foreach (KeyValuePair<string, decimal> val in res)
                {
                    ws.Cell(row, 1).Value = val.Key;
                    ws.Cell(row, 2).Value = val.Value;
                    row++;
                }
                wb.SaveAs(@"C:\temp\saldoServer.xlsx");*/
            }
            #endregion

            #region 24
            else if (type == 24)
            {
                StreamWriter sw1 = new StreamWriter(@"C:\temp\6\upload61HVS.txt", false);
                StreamWriter sw2 = new StreamWriter(@"C:\temp\6\upload61Vodootv.txt", false);
                StreamWriter sw3 = new StreamWriter(@"C:\temp\6\upload60HVS.txt", false);
                StreamWriter sw4 = new StreamWriter(@"C:\temp\6\upload60GVS.txt", false);
                var wb1 = new XLWorkbook(@"C:\temp\6\_ОДН-АУК.xlsx");
                for (int i = 5; i <= 2732; i++)
                {
                    if (wb1.Worksheet(1).Row(i).Cell(3).Value.ToString() != ""
                        && (wb1.Worksheet(1).Row(i).Cell(6).Value != null || wb1.Worksheet(1).Row(i).Cell(5).Value != null)
                        && ((wb1.Worksheet(1).Row(i).Cell(6).Value.ToString() != "" && wb1.Worksheet(1).Row(i).Cell(6).Value.ToString() != "0")
                        || (wb1.Worksheet(1).Row(i).Cell(5).Value.ToString() != "" && wb1.Worksheet(1).Row(i).Cell(5).Value.ToString() != "0")))
                    {
                        string row = "";
                        
                        row += wb1.Worksheet(1).Row(i).Cell(9).Value.ToString() + "|";
                        
                        if (wb1.Worksheet(1).Row(i).Cell(6).Value.ToString() != "" && wb1.Worksheet(1).Row(i).Cell(6).Value.ToString() != "0")
                            row += Math.Round(Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(6).Value.ToString()), 2) + "|";
                        else
                            row += "0|";
                        if (wb1.Worksheet(1).Row(i).Cell(5).Value.ToString() != "" && wb1.Worksheet(1).Row(i).Cell(5).Value.ToString() != "0")
                            row += Math.Round(Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(5).Value.ToString()), 2) + "|";
                        else
                            row += "0|";
                        sw1.WriteLine(row);
                    }
                    if (wb1.Worksheet(1).Row(i).Cell(3).Value.ToString() != ""
                       && (wb1.Worksheet(1).Row(i).Cell(7).Value != null || wb1.Worksheet(1).Row(i).Cell(8).Value != null)
                       && ((wb1.Worksheet(1).Row(i).Cell(7).Value.ToString() != "" && wb1.Worksheet(1).Row(i).Cell(7).Value.ToString() != "0")
                       || (wb1.Worksheet(1).Row(i).Cell(8).Value.ToString() != "" && wb1.Worksheet(1).Row(i).Cell(8).Value.ToString() != "0")))
                    {
                        string row2 = "";
                        row2 += wb1.Worksheet(1).Row(i).Cell(9).Value.ToString() + "|";
                        if (wb1.Worksheet(1).Row(i).Cell(7).Value.ToString() != "" && wb1.Worksheet(1).Row(i).Cell(7).Value.ToString() != "0")
                            row2 += Math.Round(Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(7).Value.ToString()), 2) + "|";
                        else
                            row2 += "0|";
                        if (wb1.Worksheet(1).Row(i).Cell(8).Value.ToString() != "" && wb1.Worksheet(1).Row(i).Cell(8).Value.ToString() != "0")
                            row2 += Math.Round(Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(8).Value.ToString()), 2) + "|";
                        else
                            row2 += "0|";
                        
                        sw2.WriteLine(row2);
                    }
                }
                var wb2 = new XLWorkbook(@"C:\temp\6\_ОДН-ПЖРТ.xlsx");
                for (int i = 5; i <= 1956; i++)
                {
                    if (wb2.Worksheet(1).Row(i).Cell(3).Value.ToString() != ""
                       && (wb2.Worksheet(1).Row(i).Cell(6).Value != null || wb2.Worksheet(1).Row(i).Cell(5).Value != null)
                       && ((wb2.Worksheet(1).Row(i).Cell(6).Value.ToString() != "" && wb2.Worksheet(1).Row(i).Cell(6).Value.ToString() != "0")
                       || (wb2.Worksheet(1).Row(i).Cell(5).Value.ToString() != "" && wb2.Worksheet(1).Row(i).Cell(5).Value.ToString() != "0")))
                    {
                        string row = "";
                        row += wb2.Worksheet(1).Row(i).Cell(7).Value.ToString() + "|";
                        if (wb2.Worksheet(1).Row(i).Cell(6).Value.ToString() != "" && wb2.Worksheet(1).Row(i).Cell(6).Value.ToString() != "0")
                            row += Math.Round(Convert.ToDecimal(wb2.Worksheet(1).Row(i).Cell(6).Value.ToString()), 2) + "|";
                        else
                            row += "0|";
                        if (wb2.Worksheet(1).Row(i).Cell(5).Value.ToString() != "" && wb2.Worksheet(1).Row(i).Cell(5).Value.ToString() != "0")
                            row += Math.Round(Convert.ToDecimal(wb2.Worksheet(1).Row(i).Cell(5).Value.ToString()), 2) + "|";
                        else
                            row += "0|";
                        sw1.WriteLine(row);
                    }
                }
                var wb3 = new XLWorkbook(@"C:\temp\6\_ОДН-ЖКС.xlsx");
                for (int i = 5; i <= 49770; i++)
                {
                    if (wb3.Worksheet(1).Row(i).Cell(3).Value.ToString() != ""
                        && (wb3.Worksheet(1).Row(i).Cell(6).Value != null || wb3.Worksheet(1).Row(i).Cell(5).Value != null)
                        && ((wb3.Worksheet(1).Row(i).Cell(6).Value.ToString() != "" && wb3.Worksheet(1).Row(i).Cell(6).Value.ToString() != "0")
                        || (wb3.Worksheet(1).Row(i).Cell(5).Value.ToString() != "" && wb3.Worksheet(1).Row(i).Cell(5).Value.ToString() != "0")))
                    {
                        string row = "";
                        
                        row += wb3.Worksheet(1).Row(i).Cell(9).Value.ToString() + "|";
                        
                        if (wb3.Worksheet(1).Row(i).Cell(5).Value.ToString() != "" && wb3.Worksheet(1).Row(i).Cell(5).Value.ToString() != "0")
                            row += Math.Round(Convert.ToDecimal(wb3.Worksheet(1).Row(i).Cell(5).Value.ToString()), 2) + "|";
                        else
                            row += "0|";
                        if (wb3.Worksheet(1).Row(i).Cell(6).Value.ToString() != "" && wb3.Worksheet(1).Row(i).Cell(6).Value.ToString() != "0")
                            row += Math.Round(Convert.ToDecimal(wb3.Worksheet(1).Row(i).Cell(6).Value.ToString()), 2) + "|";
                        else
                            row += "0|";
                        sw3.WriteLine(row);
                    }


                    if (wb3.Worksheet(1).Row(i).Cell(3).Value.ToString() != ""
                        && (wb3.Worksheet(1).Row(i).Cell(8).Value != null || wb3.Worksheet(1).Row(i).Cell(7).Value != null)
                        && ((wb3.Worksheet(1).Row(i).Cell(8).Value.ToString() != "" && wb3.Worksheet(1).Row(i).Cell(8).Value.ToString() != "0")
                        || (wb3.Worksheet(1).Row(i).Cell(7).Value.ToString() != "" && wb3.Worksheet(1).Row(i).Cell(7).Value.ToString() != "0")))
                    {
                        string row2 = "";
                        row2 += wb3.Worksheet(1).Row(i).Cell(9).Value.ToString() + "|";
                        if (wb3.Worksheet(1).Row(i).Cell(7).Value.ToString() != "" && wb3.Worksheet(1).Row(i).Cell(7).Value.ToString() != "0")
                            row2 += Math.Round(Convert.ToDecimal(wb3.Worksheet(1).Row(i).Cell(7).Value.ToString()), 2) + "|";
                        else
                            row2 += "0|";
                        if (wb3.Worksheet(1).Row(i).Cell(8).Value.ToString() != "" && wb3.Worksheet(1).Row(i).Cell(8).Value.ToString() != "0")
                            row2 += Math.Round(Convert.ToDecimal(wb3.Worksheet(1).Row(i).Cell(8).Value.ToString()), 2) + "|";
                        else
                            row2 += "0|";
                        
                        sw4.WriteLine(row2);
                    }
                }


                sw1.Close();
                sw2.Close();
                sw3.Close();
                sw4.Close();
            }
            #endregion

            #region 25
            else if (type == 25)
            {
                Dictionary<string, string> bases = new Dictionary<string, string>();
                bases.Add(@"g34", "Baza");
                bases.Add(@"g35", "Baza");
                bases.Add(@"g36", "Baza");
                bases.Add(@"g37", "Baza");
                bases.Add(@"g38", "Baza");
                bases.Add(@"g39", "Baza");
                bases.Add(@"g42", "Baza");
                bases.Add(@"g43", "Baza");
                bases.Add(@"g46", "Baza");
                bases.Add(@"g50", "Baza");
                bases.Add(@"g59", "Baza");
                bases.Add(@"g66", "Baza");
                bases.Add(@"g67", "Baza");
                bases.Add(@"g68", "Baza");
                bases.Add(@"g69", "Baza");
                bases.Add(@"g72", "Baza");
                bases.Add(@"g73", "Baza");
                bases.Add(@"g74", "Baza");
                decimal sum = 0;
                foreach (KeyValuePair<string, string> db in bases)
                {
                    string name = db.Key;
                    string dbPath = @"C:\imp\baz_gks\" + name + @"\";
                    DataTable dtHouse = dbf.SelectKartVU(dbPath);
                    for (int i = 0; i < dtHouse.Rows.Count; i++)
                    {
                        DataRow dr = dtHouse.Rows[i];
                        if(dtHouse.Rows[i][1] != null && dtHouse.Rows[i][1].ToString() != "")
                            sum += Convert.ToDecimal(dtHouse.Rows[i][1].ToString());
                    }
                }

                Console.WriteLine(sum);

                /* var wb = new XLWorkbook();
                 var ws = wb.Worksheets.Add("1");
                 int row = 2;
                 res = SortMyDictionaryByKey(res);
                 foreach (KeyValuePair<string, decimal> val in res)
                 {
                     ws.Cell(row, 1).Value = val.Key;
                     ws.Cell(row, 2).Value = val.Value;
                     row++;
                 }
                 wb.SaveAs(@"C:\temp\saldoServer.xlsx");*/
            }
            #endregion

            #region 26
            else if (type == 26)
            {
                List<string> bases = new List<string>();
                bases.Add(@"C:\imp\Паспортистка\P01");
                string dbPath = @"C:\imp\Паспортистка\P01\";
                DataTable dtHouse = dbf.SelectKGLC(dbPath);
                DataTable dtPasskart = dbf.SelectPasskart(dbPath);
                List<string> kkods = new List<string>();
                List<string> geus = new List<string>();
                geus.Add("35");
                geus.Add("36");
                geus.Add("37");
                geus.Add("38");
                geus.Add("39");
                geus.Add("42");
                geus.Add("43");
                geus.Add("46");
                geus.Add("50");
                geus.Add("59");
                geus.Add("66");
                geus.Add("67");
                geus.Add("68");
                geus.Add("69");
                geus.Add("72");
                geus.Add("73");
                geus.Add("74");
                geus.Add("92");
                for (int i = 0; i < dtHouse.Rows.Count; i++)
                {
                    if (!kkods.Contains(dtHouse.Rows[i][0].ToString()))
                    {
                        kkods.Add(dtHouse.Rows[i][0].ToString());
                        for (int j = 0; j < dtPasskart.Rows.Count; j++)
                        {
                            if (dtHouse.Rows[i][0].ToString() == dtPasskart.Rows[j][1].ToString())
                            {
                                if (!geus.Contains(dtPasskart.Rows[j][0].ToString()))
                                {
                                    dbf.DeleteKglc(dbPath, dtHouse.Rows[i][0].ToString());
                                }
                                break;
                            }
                        }
                    }
                }
            }
            #endregion

            #region 4
            else if (type == 4)
            {
                Dictionary<string, string> bases = new Dictionary<string, string>();
                bases.Add("BAZ_ALT", "78");
                bases.Add("BAZ_AYK", "79");
                bases.Add("BAZ_TREST", "82");
                bases.Add("baz_gks", "131");
                bases.Add("baz_rem", "130");
                List<string> kHouse = new List<string>();
                List<string> vu = new List<string>() { "0201", "0202" };
                var wb = new XLWorkbook();
                foreach (KeyValuePair<string, string> db in bases)
                {
                    int rows = 1;
                    string name = db.Key;
                    var ws = wb.Worksheets.Add(name);
                    string dbPath = @"C:\imp\" + name + @"\spr\";
                    string kodYK = "1111";
                    DataTable dtNormTemp = dbf.SelectNorm(dbPath);
                    DataTable dtHouseTemp = dbf.SelectHouse2(dbPath);
                    DataTable dtTarifTemp = dbf.SelectTarif(dbPath);
                    DataTable dtNorm = new DataTable();
                    DataTable dtHouse = new DataTable();
                    DataTable dtTarif = new DataTable();
                    dtNorm.Columns.Add("1");
                    dtNorm.Columns.Add("2");
                    dtNorm.Columns.Add("3");
                    dtNorm.Columns.Add("4");
                    dtHouse.Columns.Add("1");
                    dtHouse.Columns.Add("2");
                    dtHouse.Columns.Add("3");
                    dtHouse.Columns.Add("4");
                    dtHouse.Columns.Add("5");
                    dtTarif.Columns.Add("1");
                    dtTarif.Columns.Add("2");
                    dtTarif.Columns.Add("3");
                    DataRow row2;
                    DataRow row3;
                    DataRow row4;
                    for (int i = 0; i < dtNormTemp.Rows.Count; i++)
                    {
                        if (vu.Contains(dtNormTemp.Rows[i][0].ToString()))
                        {
                            row2 = dtNorm.NewRow();
                            row2["1"] = dtNormTemp.Rows[i][0].ToString();
                            row2["2"] = dtNormTemp.Rows[i][1].ToString();
                            row2["3"] = dtNormTemp.Rows[i][2].ToString();
                            row2["4"] = dtNormTemp.Rows[i][3].ToString();
                            dtNorm.Rows.Add(row2);
                        }
                    }
                    for (int i = 0; i < dtHouseTemp.Rows.Count; i++)
                    {
                        if (vu.Contains(dtHouseTemp.Rows[i][1].ToString()))
                        {
                            row3 = dtHouse.NewRow();
                            if (dtHouseTemp.Rows[i][0].ToString() == "480")
                            {
                                string str = "222";
                                str = str.Substring(1);
                            }
                            row3["1"] = dtHouseTemp.Rows[i][0].ToString();
                            row3["2"] = dtHouseTemp.Rows[i][1].ToString();
                            row3["3"] = dtHouseTemp.Rows[i][2].ToString();
                            row3["4"] = dtHouseTemp.Rows[i][3].ToString();
                            row3["5"] = dtHouseTemp.Rows[i][4].ToString();
                            dtHouse.Rows.Add(row3);
                        }
                    }
                    for (int i = 0; i < dtTarifTemp.Rows.Count; i++)
                    {
                        if (vu.Contains(dtTarifTemp.Rows[i][0].ToString()))
                        {
                            row4 = dtTarif.NewRow();
                            row4["1"] = dtTarifTemp.Rows[i][0].ToString();
                            row4["2"] = dtTarifTemp.Rows[i][1].ToString();
                            row4["3"] = dtTarifTemp.Rows[i][2].ToString();
                            dtTarif.Rows.Add(row4);
                        }
                    }
                    if (name == "baz_rem")
                    {
                        string str = "222";
                        str = str.Substring(1);
                    }
                    for (int i = 0; i < dtHouse.Rows.Count; i++)
                    {
                        if (!kHouse.Contains(dtHouse.Rows[i][0].ToString() + "|" + name))
                        {
                            if (dtHouse.Rows[i][0].ToString() == "931")
                            {
                                string sre = "etertertre";
                                sre.Substring(4);
                            }
                            if (dtHouse.Rows[i][2] != null && dtHouse.Rows[i][2].ToString() != "")
                            {
                                for (int j = 0; j < dtTarif.Rows.Count; j++)
                                {
                                    if (dtHouse.Rows[i][1].ToString() == dtTarif.Rows[j][0].ToString() && dtHouse.Rows[i][2].ToString() == dtTarif.Rows[j][1].ToString())
                                    {
                                        for (int k = 0; k < dtNorm.Rows.Count; k++)
                                        {
                                            if (dtNorm.Rows[k][0].ToString() == dtTarif.Rows[j][0].ToString() && dtNorm.Rows[k][2].ToString() == dtTarif.Rows[j][2].ToString())
                                            {
                                                ws.Cell(rows, 1).Value = dtHouse.Rows[i][0].ToString();
                                                string str = "";
                                                str = Convert2(dtNorm.Rows[k][1].ToString(), Encoding.GetEncoding(1251), Encoding.Default);
                                                ws.Cell(rows, 2).Value = str;
                                                //ws.Cell(i + 1, 2).Value = dtNormTemp.Rows[i][1].ToString();
                                                ws.Cell(rows, 3).Value = dtNorm.Rows[k][3].ToString();
                                                if (dtHouse.Rows[i][0].ToString() == "931")
                                                {
                                                    string sre = "etertertre";
                                                    sre.Substring(4);
                                                }
                                                kHouse.Add(dtHouse.Rows[i][0].ToString() + "|" + name);
                                                rows++;
                                                break;
                                            }
                                        }
                                        break;
                                    }
                                }
                            }
                            else if (dtHouse.Rows[i][3] != null && dtHouse.Rows[i][3].ToString() != "")
                            {
                                for (int j = 0; j < dtTarif.Rows.Count; j++)
                                {
                                    if (dtHouse.Rows[i][1].ToString() == dtTarif.Rows[j][0].ToString() && dtHouse.Rows[i][3].ToString() == dtTarif.Rows[j][1].ToString())
                                    {
                                        for (int k = 0; k < dtNorm.Rows.Count; k++)
                                        {
                                            if (dtNorm.Rows[k][0].ToString() == dtTarif.Rows[j][0].ToString() && dtNorm.Rows[k][2].ToString() == dtTarif.Rows[j][2].ToString())
                                            {
                                                ws.Cell(rows, 1).Value = dtHouse.Rows[i][0].ToString();
                                                string str = "";
                                                str = Convert2(dtNorm.Rows[k][1].ToString(), Encoding.GetEncoding(1251), Encoding.Default);
                                                ws.Cell(rows, 2).Value = str;
                                                //ws.Cell(i + 1, 2).Value = dtNormTemp.Rows[i][1].ToString();
                                                ws.Cell(rows, 3).Value = dtNorm.Rows[k][3].ToString();
                                                kHouse.Add(dtHouse.Rows[i][0].ToString() + "|" + name);
                                                if (dtHouse.Rows[i][0].ToString() == "931")
                                                {
                                                    string sre = "etertertre";
                                                    sre.Substring(4);
                                                }
                                                rows++;
                                                break;
                                            }
                                        }
                                        break;
                                    }
                                }
                            }
                            else if (dtHouse.Rows[i][4] != null && dtHouse.Rows[i][4].ToString() != "")
                            {
                                for (int j = 0; j < dtTarif.Rows.Count; j++)
                                {
                                    if (dtHouse.Rows[i][1].ToString() == dtTarif.Rows[j][0].ToString() && dtHouse.Rows[i][4].ToString() == dtTarif.Rows[j][1].ToString())
                                    {
                                        for (int k = 0; k < dtNorm.Rows.Count; k++)
                                        {
                                            if (dtNorm.Rows[k][0].ToString() == dtTarif.Rows[j][0].ToString() && dtNorm.Rows[k][2].ToString() == dtTarif.Rows[j][2].ToString())
                                            {
                                                ws.Cell(rows, 1).Value = dtHouse.Rows[i][0].ToString();
                                                string str = "";
                                                str = Convert2(dtNorm.Rows[k][1].ToString(), Encoding.GetEncoding(1251), Encoding.Default);
                                                ws.Cell(rows, 2).Value = str;
                                                //ws.Cell(i + 1, 2).Value = dtNormTemp.Rows[i][1].ToString();
                                                ws.Cell(rows, 3).Value = dtNorm.Rows[k][3].ToString();
                                                kHouse.Add(dtHouse.Rows[i][0].ToString());
                                                rows++;
                                                if (dtHouse.Rows[i][0].ToString() == "931" + "|" + name)
                                                {
                                                    string sre = "etertertre";
                                                    sre.Substring(4);
                                                }
                                                break;
                                            }
                                        }
                                        break;
                                    }
                                }
                            }
                            else
                            {

                            }
                        }
                    }
                }
                wb.SaveAs(@"C:\temp\qwerty6.xlsx");
            }
            #endregion

            #region 3
            else if (type == 3)
            {
                string name = "baz_gks";
                int rows = 687;
                var wb = new XLWorkbook(@"C:\temp\" + name + ".xlsx");
                var wb2 = new XLWorkbook(@"C:\temp\" + name + "2.xlsx");
                for (int i = 1; i <= rows; i++)
                {
                    bool b = true;
                    for (int j = 1; j <= rows; j++)
                    {
                        if (Convert.ToString(wb.Worksheet(1).Row(i).Cell(1).Value).Trim() == Convert.ToString(wb2.Worksheet(1).Row(j).Cell(1).Value).Trim()
                            && Convert.ToString(wb.Worksheet(1).Row(i).Cell(2).Value).Trim() == Convert.ToString(wb2.Worksheet(1).Row(j).Cell(2).Value).Trim())
                        {
                            b = false;
                        }
                    }
                    if (b)
                    {
                        wb.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                }
                wb.Save();
            }
            #endregion
            #region 2
            else if (type == 2)
            {
                var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("1");
                string dbPath = @"C:\imp\baz_sov\SPR\";
                string name = "baz_sov";
                DataTable dtHouse = dbf.SelectHouse(dbPath);
                DataTable dtStreet = dbf.SelectStreet(dbPath);
                int row = 1;
                for (int i = 0; i < dtHouse.Rows.Count; i++)
                {
                    for (int j = 0; j < dtStreet.Rows.Count; j++)
                    {
                        if (dtStreet.Rows[j][3].ToString() == dtHouse.Rows[i][9].ToString())
                        {
                            ws.Cell(row, 1).Value = dtStreet.Rows[j][1].ToString();
                            ws.Cell(row, 2).Value = dtHouse.Rows[i][1].ToString();
                            ws.Cell(row, 3).Value = dtHouse.Rows[i][10].ToString();
                            row++;
                        }
                    }
                }
                wb.SaveAs(@"C:\temp\" + name + ".xlsx");
            }
            #endregion
            #region 1
            if (type == 1)//запись из Эксельки
            {
                string nameReg = "печерская-151";
                StreamWriter sw = new StreamWriter(@"C:\temp\" + nameReg + ".txt", false, Encoding.Unicode);
                DataTable dtHouse = new DataTable();
                DataTable dtPeople = new DataTable();
                var wb = new XLWorkbook(@"C:\temp\печерская-151.xlsx");
                DataRow row;
                DataTable dt1 = new System.Data.DataTable();
                dt1.Columns.Add("1");
                dt1.Columns.Add("2");
                dt1.Columns.Add("3");
                dt1.Columns.Add("4");
                dt1.Columns.Add("5");
                dt1.Columns.Add("6");
                dt1.Columns.Add("7");
                dt1.Columns.Add("8");
                for (int i = 2; i <= 202; i++)//34276
                {
                    row = dt1.NewRow();
                    row["1"] = Convert.ToString(wb.Worksheet(1).Row(i).Cell(3).Value);
                    row["2"] = Convert.ToString(wb.Worksheet(1).Row(i).Cell(4).Value);
                    row["3"] = Convert.ToString(wb.Worksheet(1).Row(i).Cell(7).Value);
                    row["4"] = Convert.ToString(wb.Worksheet(1).Row(i).Cell(8).Value);
                    row["5"] = Convert.ToString(wb.Worksheet(1).Row(i).Cell(9).Value);
                    row["6"] = Convert.ToString(wb.Worksheet(1).Row(i).Cell(11).Value);
                    row["7"] = Convert.ToString(wb.Worksheet(1).Row(i).Cell(12).Value);
                    if (Convert.ToString(wb.Worksheet(1).Row(i).Cell(13).Value) == "Да")
                        row["8"] = 1;
                    else
                        row["8"] = 0;
                    dt1.Rows.Add(row);
                }
                //List<string> code = new List<string>() { "6700087", "6700127", "6700064", "6700128", "6700034", "6700103", "6700078", "6700030", "6700221", "6700398", "6700397" };
                //dtHouse = ora.SelectHouse(code);
                //dtPeople = ora.SelectLN4(code);
                Dictionary<string, int> manOrgName = new Dictionary<string, int>();
                List<string> p = new List<string>();
                int manOrgCode = 3000;
                int inn = 11111111;// 11111111
                int kpp = 2222222;//2222222

                /*
                for (int j = 0; j < dt1.Rows.Count; j++)
                {
                    string rowWrite = "";
                    //1
                    rowWrite += "5|";
                    //2
                    if (!manOrgName.ContainsKey(dt1.Rows[j][3].ToString()))
                    {
                        manOrgCode++;
                        manOrgName.Add(dt1.Rows[j][3].ToString(), manOrgCode);
                        rowWrite += manOrgCode + "|";
                        //3
                        rowWrite += dt1.Rows[j][3].ToString() + "|||";
                        //6
                        inn++;
                        rowWrite += inn + "|";
                        //7
                        kpp++;
                        rowWrite += kpp + "|";
                        //8
                        rowWrite += dt1.Rows[j][5].ToString() + "|";
                        //9
                        rowWrite += dt1.Rows[j][6].ToString() + "|";
                        //10
                        rowWrite += dt1.Rows[j][7].ToString() + "||";
                        //11
                        //rowIns
                        sw.WriteLine(rowWrite);
                    }

                }
                List<string> mo = new List<string>();
                for (int j = 0; j < dt1.Rows.Count; j++)
                {
                    if (!mo.Contains(dt1.Rows[j][2].ToString()))
                    {
                        dtHouse = ora.SelectHouseCode(dt1.Rows[j][2].ToString());//Загружаем дом
                        mo.Add(dt1.Rows[j][2].ToString());
                        for (int i = 0; i < dtHouse.Rows.Count; i++)
                        {
                            string rowWrite = "";
                            //1
                            rowWrite += dtHouse.Rows[i][0].ToString() + "||";
                            //3
                            rowWrite += dtHouse.Rows[i][2].ToString() + "|";
                            //4
                            rowWrite += dtHouse.Rows[i][3].ToString() + "|";
                            //5
                            if (dtHouse.Rows[i][17].ToString().Contains("Тольятти"))
                                rowWrite += "|";
                            else
                            {
                                if (dtHouse.Rows[i][17] != null && dtHouse.Rows[i][17].ToString() != "")
                                    rowWrite += dtHouse.Rows[i][17].ToString() + "|";
                                else
                                    rowWrite += dtHouse.Rows[i][4].ToString().Split(',')[0] + "|";
                            }
                            //6
                            if (dtHouse.Rows[i][18] != null && dtHouse.Rows[i][18].ToString() != "")
                                rowWrite += dtHouse.Rows[i][18].ToString() + "|";
                            else
                                rowWrite += "|";
                            //7
                            rowWrite += dtHouse.Rows[i][19].ToString() + "|";
                            //8
                            if (dtHouse.Rows[i][20] != null && dtHouse.Rows[i][20].ToString() != "")
                                rowWrite += dtHouse.Rows[i][20].ToString() + "|";
                            else
                                rowWrite += "|";
                            //9
                            rowWrite += "666" + "|";
                            //10
                            rowWrite += dtHouse.Rows[i][6].ToString() + "|";
                            //11
                            if (dtHouse.Rows[i][7] != null && dtHouse.Rows[i][7].ToString() != "")
                                rowWrite += dtHouse.Rows[i][7].ToString() + "|";
                            else
                                rowWrite += "1|";
                            //12
                            rowWrite += dtHouse.Rows[i][8].ToString() + "|";
                            //13
                            rowWrite += dtHouse.Rows[i][9].ToString() + "|||";
                            //16
                            rowWrite += "100||";
                            //rowWrite += dtHouse.Rows[i][12].ToString() + "||";
                            //18
                            rowWrite += dtHouse.Rows[i][14].ToString() + "||";
                            bool t = true;
                            //20
                            if (dtHouse.Rows[i][16] != null && dtHouse.Rows[i][16].ToString() != "")
                                rowWrite += dtHouse.Rows[i][16].ToString() + "|";
                            else
                            {
                                rowWrite += "0000|";
                            }
                            if (t)
                                sw.WriteLine(rowWrite);
                        }
                    }
                }*/
                Dictionary<string, int> dict = new Dictionary<string, int>();
                int ownCode = 99999;

                for (int j = 0; j < dt1.Rows.Count; j++)
                {
                    //код ЛС уыеличиваем
                    
                    string rowWrite = "";
                    //1
                    rowWrite += "4||";
                    //3
                    rowWrite += "1921|";
                    //4
                    rowWrite += dt1.Rows[j][0].ToString() + "|";
                    ownCode--;
                    //5
                    rowWrite += "1|";
                    //6
                    rowWrite += dt1.Rows[j][1].ToString() + "||||";
                    //10
                    rowWrite += dt1.Rows[j][2].ToString() + "||||||";
                    //16
                    rowWrite += dt1.Rows[j][3].ToString() + "|";
                    //17
                    rowWrite += "0|";
                    //18
                    rowWrite += "0|";
                    //19
                    rowWrite += "1|";
                    //20
                    rowWrite += dt1.Rows[j][5].ToString() + "|";
                    //21
                    rowWrite += dt1.Rows[j][6].ToString() + "|||";
                    //24
                    rowWrite += "0|||||||||";
                    //33
                    rowWrite += dt1.Rows[j][7].ToString() + "||||||";
                    sw.WriteLine(rowWrite);
                }
                sw.Close();
            }
            #endregion
            #region 0
            else if (type == 0)
            {
                DataTable dtHouse = new DataTable();
                DataTable dtPeople = new DataTable();
                var wb = new XLWorkbook(@"C:\temp\Копия Ставропольский район.xlsx");
                DataRow row;
                DataTable dt1 = new System.Data.DataTable();
                dt1.Columns.Add("1");
                for (int i = 5; i <= 107; i++)//34276
                {
                    row = dt1.NewRow();
                    row["1"] = Convert.ToString(wb.Worksheet(1).Row(i).Cell(2).Value);
                    dt1.Rows.Add(row);
                }
                //List<string> code = new List<string>() { "6700087", "6700127", "6700064", "6700128", "6700034", "6700103", "6700078", "6700030", "6700221", "6700398", "6700397" };
                //dtHouse = ora.SelectHouse(code);
                //dtPeople = ora.SelectLN4(code);
                StreamWriter sw = new StreamWriter(@"C:\temp\Ставропольский.txt", false);
                for (int j = 0; j < dt1.Rows.Count; j++)
                {
                    dtHouse = ora.SelectHouseCode(dt1.Rows[j][0].ToString());


                    for (int i = 0; i < dtHouse.Rows.Count; i++)
                    {
                        string rowWrite = "";
                        rowWrite += dtHouse.Rows[i][0].ToString() + "||";
                        rowWrite += dtHouse.Rows[i][2].ToString() + "|";
                        rowWrite += dtHouse.Rows[i][3].ToString() + "|";
                        if (dtHouse.Rows[i][17].ToString().Contains("Тольятти"))
                            rowWrite += "|";
                        else
                        {
                            if (dtHouse.Rows[i][17] != null && dtHouse.Rows[i][17].ToString() != "")
                                rowWrite += dtHouse.Rows[i][17].ToString() + "|";
                            else
                                rowWrite += dtHouse.Rows[i][4].ToString().Split(',')[0] + "|";
                        }
                        if (dtHouse.Rows[i][18] != null && dtHouse.Rows[i][18].ToString() != "")
                            rowWrite += dtHouse.Rows[i][18].ToString() + "|";
                        else
                            rowWrite += "|";
                        rowWrite += dtHouse.Rows[i][19].ToString() + "|";
                        if (dtHouse.Rows[i][20] != null && dtHouse.Rows[i][20].ToString() != "")
                            rowWrite += dtHouse.Rows[i][20].ToString() + "|";
                        else
                            rowWrite += "|";
                        rowWrite += dtHouse.Rows[i][5].ToString() + "|";
                        rowWrite += dtHouse.Rows[i][6].ToString() + "|";
                        if (dtHouse.Rows[i][7] != null && dtHouse.Rows[i][7].ToString() != "")
                            rowWrite += dtHouse.Rows[i][7].ToString() + "|";
                        else
                            rowWrite += "1|";
                        rowWrite += dtHouse.Rows[i][8].ToString() + "|";
                        rowWrite += dtHouse.Rows[i][9].ToString() + "|||";
                        rowWrite += dtHouse.Rows[i][12].ToString() + "||";
                        rowWrite += dtHouse.Rows[i][14].ToString() + "||";
                        bool t = true;
                        if (dtHouse.Rows[i][16] != null && dtHouse.Rows[i][16].ToString() != "")
                            rowWrite += dtHouse.Rows[i][16].ToString() + "|";
                        else
                        {
                            rowWrite += "6302800000000|";
                        }
                        if (t)
                            sw.WriteLine(rowWrite);
                    }
                }
                Dictionary<string, int> dict = new Dictionary<string, int>();
                //int ownCode = 99999;
                for (int j = 0; j < dt1.Rows.Count; j++)
                {
                    dtPeople = ora.SelectLN4Code(dt1.Rows[j][0].ToString());
                    for (int i = 0; i < dtPeople.Rows.Count; i++)
                    {
                        if (dict.ContainsKey(dtPeople.Rows[i][1].ToString()))
                            dict[dtPeople.Rows[i][1].ToString()]++;
                        else
                            dict.Add(dtPeople.Rows[i][1].ToString(), 1);
                        string rowWrite = "";
                        rowWrite += dtPeople.Rows[i][0].ToString() + "||";
                        rowWrite += dtPeople.Rows[i][1].ToString() + "|";
                        rowWrite += dtPeople.Rows[i][1].ToString() + dict[dtPeople.Rows[i][1].ToString()].ToString().PadLeft(5, '0') + "|";
                        //ownCode--;
                        rowWrite += dtPeople.Rows[i][3].ToString() + "|";
                        rowWrite += dtPeople.Rows[i][4].ToString().Replace("|", "/").Trim() + "||||";
                        rowWrite += dtPeople.Rows[i][5].ToString().Replace("|", "/") + "||||||";
                        if (dtPeople.Rows[i][6] != null && dtPeople.Rows[i][6].ToString() != "")
                            rowWrite += dtPeople.Rows[i][6].ToString() + "|";
                        else
                            rowWrite += "0|";
                        rowWrite += dtPeople.Rows[i][7].ToString() + "|";
                        rowWrite += dtPeople.Rows[i][8].ToString() + "|";
                        rowWrite += dtPeople.Rows[i][9].ToString() + "|";
                        if (dtPeople.Rows[i][10] != null && dtPeople.Rows[i][10].ToString() != "")
                            rowWrite += dtPeople.Rows[i][10].ToString() + "|";
                        else
                            rowWrite += "0|";
                        if (dtPeople.Rows[i][11] != null && dtPeople.Rows[i][11].ToString() != "")
                            rowWrite += dtPeople.Rows[i][11].ToString() + "|||";
                        else
                            rowWrite += "|||";
                        rowWrite += dtPeople.Rows[i][12].ToString() + "|||||||||";
                        rowWrite += dtPeople.Rows[i][13].ToString() + "|||||";
                        rowWrite += "|";
                        sw.WriteLine(rowWrite);
                    }
                }
                sw.Close();
            }
            #endregion

            #region 27
            else if (type == 27)
            {
                DataTable dtHouse = new DataTable();
                DataTable dtPeople = new DataTable();
                var wb = new XLWorkbook(@"C:\temp\Копия Ставропольский район.xlsx");
                var wb2 = new XLWorkbook(@"C:\temp\Копия Электронное ЖКХ.xlsx");
                DataRow row;
                DataTable dt1 = new System.Data.DataTable();
                dt1.Columns.Add("1");
                dt1.Columns.Add("2");
                for (int i = 5; i <= 108; i++)//34276
                {
                    row = dt1.NewRow();
                    row["1"] = Convert.ToString(wb.Worksheet(1).Row(i).Cell(2).Value);
                    row["2"] = Convert.ToString(wb.Worksheet(1).Row(i).Cell(4).Value).Replace("Ставропольский р-н с. ", "").Replace("Ставропольский р-н пос. ", "").Replace(" ","");
                    dt1.Rows.Add(row);
                }
                //List<string> code = new List<string>() { "6700087", "6700127", "6700064", "6700128", "6700034", "6700103", "6700078", "6700030", "6700221", "6700398", "6700397" };
                //dtHouse = ora.SelectHouse(code);
                //dtPeople = ora.SelectLN4(code);
                StreamWriter sw = new StreamWriter(@"C:\temp\Ставропольский.txt", false);
                for (int j = 0; j < dt1.Rows.Count; j++)
                {
                    dtHouse = ora.SelectHouseCode(dt1.Rows[j][0].ToString());
                    for (int i = 0; i < dtHouse.Rows.Count; i++)
                    {
                        string rowWrite = "";
                        rowWrite += dtHouse.Rows[i][0].ToString() + "||";
                        rowWrite += dtHouse.Rows[i][2].ToString() + "|";
                        rowWrite += dtHouse.Rows[i][3].ToString() + "|";
                        if (dtHouse.Rows[i][17].ToString().Contains("Тольятти"))
                            rowWrite += "|";
                        else
                        {
                            if (dtHouse.Rows[i][17] != null && dtHouse.Rows[i][17].ToString() != "")
                                rowWrite += dtHouse.Rows[i][17].ToString() + "|";
                            else
                                rowWrite += dtHouse.Rows[i][4].ToString().Split(',')[0] + "|";
                        }
                        if (dtHouse.Rows[i][18] != null && dtHouse.Rows[i][18].ToString() != "")
                            rowWrite += dtHouse.Rows[i][18].ToString() + "|";
                        else
                            rowWrite += "|";
                        rowWrite += dtHouse.Rows[i][19].ToString() + "|";
                        if (dtHouse.Rows[i][20] != null && dtHouse.Rows[i][20].ToString() != "")
                            rowWrite += dtHouse.Rows[i][20].ToString() + "|";
                        else
                            rowWrite += "|";
                        rowWrite += dtHouse.Rows[i][5].ToString() + "|";
                        rowWrite += dtHouse.Rows[i][6].ToString() + "|";
                        if (dtHouse.Rows[i][7] != null && dtHouse.Rows[i][7].ToString() != "")
                            rowWrite += dtHouse.Rows[i][7].ToString() + "|";
                        else
                            rowWrite += "1|";
                        rowWrite += dtHouse.Rows[i][8].ToString() + "|";
                        rowWrite += dtHouse.Rows[i][9].ToString() + "|||";
                        rowWrite += dtHouse.Rows[i][12].ToString() + "||";
                        rowWrite += dtHouse.Rows[i][14].ToString() + "||";
                        bool t = true;
                        if (dtHouse.Rows[i][16] != null && dtHouse.Rows[i][16].ToString() != "")
                            rowWrite += dtHouse.Rows[i][16].ToString() + "|";
                        else
                        {
                            rowWrite += "6302800000000|";
                        }
                        if (t)
                            sw.WriteLine(rowWrite);
                    }
                }
                Dictionary<string, int> dict = new Dictionary<string, int>();
                 string[] stringSeparators = new string[] { ", кв. " };
                for (int j = 0; j < dt1.Rows.Count; j++)
                {
                    for (int i = 3; i <=1506; i++)
                    {
                        //string str = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Split(stringSeparators, StringSplitOptions.None)[0].Replace(" ", "");
                        //string str2 = dt1.Rows[j][1].ToString();
                        //Console.WriteLine(str);
                        //Console.WriteLine(str2);
                        if (dt1.Rows[j][1].ToString().ToUpper() ==
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Split(stringSeparators, StringSplitOptions.None)[0].Replace(" ", "").ToUpper())
                        {
                            wb2.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                            if (dict.ContainsKey(dt1.Rows[j][0].ToString()))
                                dict[dt1.Rows[j][0].ToString()]++;
                            else
                                dict.Add(dt1.Rows[j][0].ToString(), 1);
                            string rowWrite = "";
                            //1 Тип строки
                            rowWrite += "4||";
                            //3.Уникальный код дома в системе отправителя
                            rowWrite += dt1.Rows[j][0].ToString() + "|";
                            //4.	№ ЛС в системе поставщика
                            rowWrite += dt1.Rows[j][0].ToString() + dict[dt1.Rows[j][0].ToString()].ToString().PadLeft(5, '0') + "|";
                            //5.	Тип ЛС (1 – жилая квартира, 2 – субабонент / арендатор)
                            rowWrite += "1|";
                            //6.	Фамилия квартиросъемщика
                            rowWrite += Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Replace("|", "/").Trim() + "||||";
                            //10.	Квартира
                            rowWrite += Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Split(stringSeparators, StringSplitOptions.None)[1].Replace("|", "/").Trim() + "||||||";
                            //16.	Количество проживающих
                            rowWrite += "0|";
                            //17.	Количество врем. прибывших жильцов
                            rowWrite += "0|";
                            //18.	Количество  врем. убывших жильцов
                            rowWrite += "0|";
                            //19.	Количество комнат
                            rowWrite += "1|";
                            //20.	Общая площадь (площадь применяемая для расчета большинства площадных услуг)
                            if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value) != "не начисл.")
                                rowWrite += Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value) + "|";
                            else
                                rowWrite += "0|";
                            //21.	Жилая площадь 
                            rowWrite += "|||";
                            //24.	Признак коммунальной квартиры(1-да, 0 –нет)
                            rowWrite += "0|||||||||";
                            //33.	Наличие забора из открытой системы отопления (1-да, 0 –нет)
                            rowWrite += "1||||||";
                            sw.WriteLine(rowWrite);
                        }
                    }
                }
                sw.Close();
                wb2.Save();
            }
            #endregion

            #region 28
            else if (type == 28)
            {
                DataTable dtHouse = new DataTable();
                DataTable dtPeople = new DataTable();
                var wb = new XLWorkbook(@"C:\temp\Копия Ставропольский район.xlsx");
                var wb2 = new XLWorkbook(@"C:\temp\Копия Электронное ЖКХ.xlsx");
                DataRow row;
                DataTable dt1 = new System.Data.DataTable();
                dt1.Columns.Add("1");
                dt1.Columns.Add("2");
                for (int i = 5; i <= 108; i++)//34276
                {
                    row = dt1.NewRow();
                    row["1"] = Convert.ToString(wb.Worksheet(1).Row(i).Cell(2).Value);
                    row["2"] = Convert.ToString(wb.Worksheet(1).Row(i).Cell(4).Value).Replace("Ставропольский р-н с. ", "").Replace("Ставропольский р-н пос. ", "").Replace(" ", "");
                    dt1.Rows.Add(row);
                }

                
                string[] stringSeparators = new string[] { ", кв. " };
                for (int j = 0; j < dt1.Rows.Count; j++)
                {
                    for (int i = 3; i <= 1506; i++)
                    {
                        if (dt1.Rows[j][1].ToString().ToUpper() ==
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Split(stringSeparators, StringSplitOptions.None)[0].Replace(" ", "").ToUpper())
                        {
                            string str = ora.InsertPeople(dt1.Rows[j][0].ToString(), 
                                Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Split(stringSeparators, StringSplitOptions.None)[1].Trim(),
                                Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value), Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value));
                            if (str != "ЗАГРУЖЕНО")
                                Console.WriteLine();
                        }
                    }
                }
                wb2.Save();
            }
            #endregion

            #region 29
            else if (type == 29)
            {
                DataTable dtHouse = new DataTable();
                DataTable dtPeople = new DataTable();
                var wb2 = new XLWorkbook(@"C:\temp\Копия Список лицевых счетов.xlsx");
                string[] stringSeparators = new string[] { ", кв." };
                string[] stringSeparators2 = new string[] { "офис  " };
                for (int i = 77; i <= 115; i++)
                {
                    try
                    {
                        string str = ora.InsertPeople2("8901531",
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim().Split(stringSeparators, StringSplitOptions.None)[1],
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim(), Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value)
                            , Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value));
                        if (str != "ЗАГРУЖЕНО")
                            Console.WriteLine(str);
                    }
                    catch
                    {
                        try
                        {
                            string str = ora.InsertOffice("8901401",
                            "офис " + Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim().Split(stringSeparators2, StringSplitOptions.None)[1],
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim(), Convert.ToString(wb2.Worksheet(1).Row(i).Cell(9).Value));
                            if (str != "ЗАГРУЖЕНО")
                                Console.WriteLine(str);
                        }
                        catch
                        {
                        }
                    }

                }
                wb2.Save();
            }
            #endregion

            #region 30
            else if (type == 30)
            {
                var wb2 = new XLWorkbook(@"C:\temp\МУП УД.xlsx");
                for (int i = 4; i <= 108; i++)
                {

                    string str = ora.InsertHouseManOrg(
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim(),
                        "02.03.2015", "17729406");
                    if (str != "ЗАГРУЖЕНО")
                        Console.WriteLine(str);
                }
                wb2.Save();
            }
            #endregion

            #region 310
            else if (type == 310)
            {
                var wb2 = new XLWorkbook(@"C:\temp\forImport.xlsx");
                for (int i = 3; i <= 9; i++)
                {
                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value).Trim() != "")
                    {
                        string str = ora.InsertHouseManOrg(
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim().Substring(0,10),
                        "17731219");
                        if (str != "ЗАГРУЖЕНО")
                            Console.WriteLine(str + "||1||" + i.ToString());
                    }
                }
                for (int i = 4; i <= 6; i++)
                {
                    if (Convert.ToString(wb2.Worksheet(2).Row(i).Cell(5).Value).Trim() != "")
                    {
                        string str = ora.InsertHouseManOrg(
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
                        string str = ora.InsertHouseManOrg(
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
                        string str = ora.InsertHouseManOrg(
                        Convert.ToString(wb2.Worksheet(4).Row(i).Cell(5).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(4).Row(i).Cell(4).Value).Trim().Substring(0,10),
                        "17731219");
                        if (str != "ЗАГРУЖЕНО")
                            Console.WriteLine(str + "||4||" + i.ToString());
                    }
                }
                wb2.Save();
            }
            #endregion

            #region 300
            else if (type == 300)
            {
                var wb2 = new XLWorkbook(@"C:\temp\forImport.xlsx");
                for (int i = 3; i <= 9; i++)
                {

                    DataTable dt = ora.SelectGkhCode(
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim());
                    if (dt.Rows.Count == 1)
                    {
                        wb2.Worksheet(1).Row(i).Cell(5).Value = dt.Rows[0][0];
                    }
                    else if (dt.Rows.Count == 0)
                    {
                        wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Red;
                    }
                    else if (dt.Rows.Count >= 2)
                    {
                        wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                }
                for (int i = 4; i <= 6; i++)
                {

                    DataTable dt = ora.SelectGkhCode(
                        Convert.ToString(wb2.Worksheet(2).Row(i).Cell(2).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(2).Row(i).Cell(3).Value).Trim());
                    if (dt.Rows.Count == 1)
                    {
                        wb2.Worksheet(2).Row(i).Cell(5).Value = dt.Rows[0][0];
                    }
                    else if (dt.Rows.Count == 0)
                    {
                        wb2.Worksheet(2).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Red;
                    }
                    else if (dt.Rows.Count >= 2)
                    {
                        wb2.Worksheet(2).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                }
                for (int i = 4; i <= 32; i++)
                {

                    DataTable dt = ora.SelectGkhCode(
                        Convert.ToString(wb2.Worksheet(3).Row(i).Cell(2).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(3).Row(i).Cell(3).Value).Trim());
                    if (dt.Rows.Count == 1)
                    {
                        wb2.Worksheet(3).Row(i).Cell(5).Value = dt.Rows[0][0];
                    }
                    else if (dt.Rows.Count == 0)
                    {
                        wb2.Worksheet(3).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Red;
                    }
                    else if (dt.Rows.Count >= 2)
                    {
                        wb2.Worksheet(3).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                }
                for (int i = 3; i <= 68; i++)
                {

                    DataTable dt = ora.SelectGkhCode(
                        Convert.ToString(wb2.Worksheet(4).Row(i).Cell(2).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(4).Row(i).Cell(3).Value).Trim());
                    if (dt.Rows.Count == 1)
                    {
                        wb2.Worksheet(4).Row(i).Cell(5).Value = dt.Rows[0][0];
                    }
                    else if (dt.Rows.Count == 0)
                    {
                        wb2.Worksheet(4).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Red;
                    }
                    else if (dt.Rows.Count >= 2)
                    {
                        wb2.Worksheet(4).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                }
                wb2.Save();
            }
            #endregion

            //Поставщики коммунальных услуг
            #region 31
            else if (type == 31)
            {

                string str = ora.InsertCommunalOrg("9548129");
                if (str != "ЗАГРУЖЕНО")
                    Console.WriteLine(str);
            }
            #endregion

            #region 42
            else if (type == 42)
            {

                string str = ora.InsertResOrg("18512502");
                if (str != "ЗАГРУЖЕНО")
                    Console.WriteLine(str);
            }
            #endregion

            #region 32
            else if (type == 32)
            {
                var wb2 = new XLWorkbook(@"C:\temp\Сведения о квартирах.xlsx");
                for (int i = 2; i <= 256; i++)
                {
                    try
                    {
                        string str = pg.InsertPeople5("9100451",
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
            #endregion

            #region 33
            else if (type == 33)
            {
                DataTable dtHouse = new DataTable();
                DataTable dtPeople = new DataTable();
                var wb2 = new XLWorkbook(@"C:\temp\Копия Реестр исходных данных кап.ремонт.xlsx");
                string[] stringSeparators = new string[] { ", кв." };
                for (int i = 8; i <= 111; i++)
                {
                    try
                    {
                        string str = ora.InsertPeople("8800162",
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
            #endregion

            #region 34
            else if (type == 34)
            {
                DataTable dtHouse = new DataTable();
                DataTable dtPeople = new DataTable();
                var wb2 = new XLWorkbook(@"C:\temp\Копия Площадь 76.xlsx");
                for (int i = 3; i <= 441; i++)
                {
                    try
                    {
                        string str = ora.InsertPeople("9700035",
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
            #endregion

            #region 35
            else if (type == 35)
            {
                DataTable dtHouse = new DataTable();
                DataTable dtPeople = new DataTable();
                var wb2 = new XLWorkbook(@"C:\temp\Копия Фрунз. 8 В.xlsx");
                for (int i = 6; i <= 17; i++)
                {
                    try
                    {
                        string str = ora.InsertPeople("9700760",
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
            #endregion

            #region 36
            else if (type == 36)
            {
                DataTable dtHouse = new DataTable();
                DataTable dtPeople = new DataTable();
                var wb2 = new XLWorkbook(@"C:\temp\Копия Ворош.15 жильцы.xlsx");
                for (int i = 6; i <= 293; i++)
                {
                    try
                    {
                        string str = ora.InsertPeople("9700352",
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim(),
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim(),
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value),
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value));
                        if (str != "ЗАГРУЖЕНО")
                            Console.WriteLine(str);
                    }
                    catch
                    {

                    }

                }
                wb2.Save();
            }
            #endregion

            //Собственники с пробегом по всем файлам
            #region 37
            else if (type == 37)
            {
                DirectoryInfo dir = new DirectoryInfo(@"C:\temp\houses4");
                string[] stringSeparators = new string[] { "Кв." };
                foreach (var item in dir.GetFiles())
                {
                    Console.WriteLine(item.Name);
                    var wb2 = new XLWorkbook(@"C:\temp\houses4\" + item.Name);
                    for (int i = 11; i <= 1000; i++)
                    {
                        if (wb2.Worksheet(1).Row(i).Cell(1).Value == null || Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim() == "")
                            break;
                        try
                        {
                            string str = ora.InsertPeople4(item.Name.Substring(0, item.Name.Length - 5),
                                Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim().Split(stringSeparators, StringSplitOptions.None)[1].Trim(),
                                "",
                                Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value),
                                "да",
                                Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value));
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
            #endregion

            #region 38
            else if (type == 38)
            {
                var wb2 = new XLWorkbook(@"C:\temp\Свед. по квартирам.xlsx");
                for (int i = 2; i <= 209; i++)
                {
                    try
                    {
                        string str = ora.InsertPeople2("9700521",
                            Convert.ToInt32(wb2.Worksheet(1).Row(i).Cell(1).Value),
                            "",
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value),
                            "0",
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value),
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value));
                        if (str != "ЗАГРУЖЕНО")
                            Console.WriteLine(str);
                    }
                    catch
                    {

                    }

                }
                wb2.Save();
            }
            #endregion

            #region 39
            else if (type == 39)
            {
                var wb2 = new XLWorkbook(@"C:\temp\Сведения о квартирах.xlsx");
                for (int i = 8; i <= 15; i++)
                {
                    try
                    {
                        string str = ora.InsertPeople("8000058",
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim(),
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim(),
                            "0",
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value), 
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value),
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value));
                        if (str != "ЗАГРУЖЕНО")
                            Console.WriteLine(str);
                    }
                    catch
                    {

                    }

                }
                wb2.Save();
            }
            #endregion

            #region 40
            else if (type == 40)
            {
                DataTable dt = ora.SelectDubl();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ora.DelDubl(dt.Rows[i]);
                }
            }
            #endregion

            #region 41 Формирование Excel файла по текущему ремонту
            else if (type == 41)
            {
                string inn = "";
                Console.Write("Введите инн управляющей организации:");
                inn = Console.ReadLine();
                string year = "";
                Console.Write("Введите год:");
                year = Console.ReadLine();
                DataTable obj = ora.SelectCurRepair(inn, year);
                var wb = new XLWorkbook();
                wb.AddWorksheet("тек.ремонт");
                wb.Worksheet(1).Range("A1", "A2").Merge();
                wb.Worksheet(1).Column(1).Width = 15;
                wb.Worksheet(1).Row(1).Cell(1).Value = "Код дома";
                wb.Worksheet(1).Row(1).Cell(1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wb.Worksheet(1).Range("B1", "B2").Merge();
                wb.Worksheet(1).Column(2).Width = 35;
                wb.Worksheet(1).Row(1).Cell(2).Value = "Адрес";
                wb.Worksheet(1).Row(1).Cell(2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wb.Worksheet(1).Range("C1", "C2").Merge();
                wb.Worksheet(1).Column(3).Width = 47;
                wb.Worksheet(1).Row(1).Cell(3).Value = "Вид работы";
                wb.Worksheet(1).Row(1).Cell(3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wb.Worksheet(1).Range("D1", "F1").Merge();
                wb.Worksheet(1).Column(4).Width = 15;
                wb.Worksheet(1).Row(1).Cell(4).Value = "План";
                wb.Worksheet(1).Row(1).Cell(4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wb.Worksheet(1).Row(2).Cell(4).Value = "Месяц";
                wb.Worksheet(1).Row(2).Cell(4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wb.Worksheet(1).Column(5).Width = 15;
                wb.Worksheet(1).Row(2).Cell(5).Value = "Сумма";
                wb.Worksheet(1).Row(2).Cell(5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wb.Worksheet(1).Column(6).Width = 15;
                wb.Worksheet(1).Row(2).Cell(6).Value = "Объем";
                wb.Worksheet(1).Row(2).Cell(6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wb.Worksheet(1).Range("G1", "I1").Merge();
                wb.Worksheet(1).Column(7).Width = 15;
                wb.Worksheet(1).Row(1).Cell(7).Value = "Факт";
                wb.Worksheet(1).Row(1).Cell(7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wb.Worksheet(1).Row(2).Cell(7).Value = "Месяц";
                wb.Worksheet(1).Row(2).Cell(7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wb.Worksheet(1).Column(8).Width = 15;
                wb.Worksheet(1).Row(2).Cell(8).Value = "Сумма";
                wb.Worksheet(1).Row(2).Cell(8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                wb.Worksheet(1).Column(9).Width = 15;
                wb.Worksheet(1).Row(2).Cell(9).Value = "Объем";
                wb.Worksheet(1).Row(2).Cell(9).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                string address = "";
                string typeWork = "";
                int rowMove = 2;
                for (int i = 0; i < obj.Rows.Count; i++)
                {
                    rowMove++;
                    if (address != obj.Rows[i][1].ToString())
                    {
                        address = obj.Rows[i][1].ToString();
                        typeWork = "";
                        if (obj.Rows[i][0] != null)
                            wb.Worksheet(1).Row(rowMove).Cell(1).Value = obj.Rows[i][0].ToString();
                        else
                            wb.Worksheet(1).Row(rowMove).Cell(1).Value = "-";
                        if (obj.Rows[i][1] != null)
                            wb.Worksheet(1).Row(rowMove).Cell(2).Value = obj.Rows[i][1].ToString();
                        else
                            wb.Worksheet(1).Row(rowMove).Cell(2).Value = "-";
                    }
                    if (typeWork != obj.Rows[i][2].ToString())
                    {
                        typeWork = obj.Rows[i][2].ToString();
                        if (obj.Rows[i][2] != null)
                            wb.Worksheet(1).Row(rowMove).Cell(3).Value = obj.Rows[i][2].ToString();
                        else
                            wb.Worksheet(1).Row(rowMove).Cell(3).Value = "-";
                    }
                    if (obj.Rows[i][3] != null && obj.Rows[i][3].ToString() != "")
                    {
                        switch (obj.Rows[i][3].ToString().Split('.')[1])
                        {
                            case "01":
                            {
                                wb.Worksheet(1).Row(rowMove).Cell(4).Value = "Январь";
                                break;
                            }
                            case "02":
                            {
                                wb.Worksheet(1).Row(rowMove).Cell(4).Value = "Февраль";
                                break;
                            }
                            case "03":
                            {
                                wb.Worksheet(1).Row(rowMove).Cell(4).Value = "Март";
                                break;
                            }
                            case "04":
                            {
                                wb.Worksheet(1).Row(rowMove).Cell(4).Value = "Апрель";
                                break;
                            }
                            case "05":
                            {
                                wb.Worksheet(1).Row(rowMove).Cell(4).Value = "Май";
                                break;
                            }
                            case "06":
                            {
                                wb.Worksheet(1).Row(rowMove).Cell(4).Value = "Июнь";
                                break;
                            }
                            case "07":
                            {
                                wb.Worksheet(1).Row(rowMove).Cell(4).Value = "Июль";
                                break;
                            }
                            case "08":
                            {
                                wb.Worksheet(1).Row(rowMove).Cell(4).Value = "Август";
                                break;
                            }
                            case "09":
                            {
                                wb.Worksheet(1).Row(rowMove).Cell(4).Value = "Сентябрь";
                                break;
                            }
                            case "10":
                            {
                                wb.Worksheet(1).Row(rowMove).Cell(4).Value = "Октябрь";
                                break;
                            }
                            case "11":
                            {
                                wb.Worksheet(1).Row(rowMove).Cell(4).Value = "Ноябрь";
                                break;
                            }
                            case "12":
                            {
                                wb.Worksheet(1).Row(rowMove).Cell(4).Value = "Декабрь";
                                break;
                            }
                        }
                    }
                    else
                        wb.Worksheet(1).Row(rowMove).Cell(4).Value = "-";
                    if (obj.Rows[i][4] != null)
                        wb.Worksheet(1).Row(rowMove).Cell(5).Value = obj.Rows[i][4].ToString();
                    else
                        wb.Worksheet(1).Row(rowMove).Cell(5).Value = "-";
                    if (obj.Rows[i][5] != null)
                        wb.Worksheet(1).Row(rowMove).Cell(6).Value = obj.Rows[i][5].ToString();
                    else
                        wb.Worksheet(1).Row(rowMove).Cell(6).Value = "-";
                    if (obj.Rows[i][6] != null && obj.Rows[i][6].ToString() != "")
                    {
                        switch (obj.Rows[i][6].ToString().Split('.')[1])
                        {
                            case "01":
                                {
                                    wb.Worksheet(1).Row(rowMove).Cell(7).Value = "Январь";
                                    break;
                                }
                            case "02":
                                {
                                    wb.Worksheet(1).Row(rowMove).Cell(7).Value = "Февраль";
                                    break;
                                }
                            case "03":
                                {
                                    wb.Worksheet(1).Row(rowMove).Cell(7).Value = "Март";
                                    break;
                                }
                            case "04":
                                {
                                    wb.Worksheet(1).Row(rowMove).Cell(7).Value = "Апрель";
                                    break;
                                }
                            case "05":
                                {
                                    wb.Worksheet(1).Row(rowMove).Cell(7).Value = "Май";
                                    break;
                                }
                            case "06":
                                {
                                    wb.Worksheet(1).Row(rowMove).Cell(7).Value = "Июнь";
                                    break;
                                }
                            case "07":
                                {
                                    wb.Worksheet(1).Row(rowMove).Cell(7).Value = "Июль";
                                    break;
                                }
                            case "08":
                                {
                                    wb.Worksheet(1).Row(rowMove).Cell(7).Value = "Август";
                                    break;
                                }
                            case "09":
                                {
                                    wb.Worksheet(1).Row(rowMove).Cell(7).Value = "Сентябрь";
                                    break;
                                }
                            case "10":
                                {
                                    wb.Worksheet(1).Row(rowMove).Cell(7).Value = "Октябрь";
                                    break;
                                }
                            case "11":
                                {
                                    wb.Worksheet(1).Row(rowMove).Cell(7).Value = "Ноябрь";
                                    break;
                                }
                            case "12":
                                {
                                    wb.Worksheet(1).Row(rowMove).Cell(7).Value = "Декабрь";
                                    break;
                                }
                        }
                    }
                    else
                        wb.Worksheet(1).Row(rowMove).Cell(7).Value = "-";
                    if (obj.Rows[i][7] != null)
                        wb.Worksheet(1).Row(rowMove).Cell(8).Value = obj.Rows[i][7].ToString();
                    else
                        wb.Worksheet(1).Row(rowMove).Cell(8).Value = "-";
                    if (obj.Rows[i][8] != null)
                        wb.Worksheet(1).Row(rowMove).Cell(9).Value = obj.Rows[i][8].ToString();
                    else
                        wb.Worksheet(1).Row(rowMove).Cell(9).Value = "-";
                }
                wb.Worksheet(1).Range("A1", "I" + rowMove).Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                wb.Worksheet(1).Range("A1", "I" + rowMove).Style.Border.RightBorder = XLBorderStyleValues.Thin;
                wb.Worksheet(1).Range("A1", "I" + rowMove).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                wb.Worksheet(1).Range("A1", "I" + rowMove).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                wb.SaveAs(@"C:\temp\temp11111.xlsx");
            }
            #endregion

            #region 43 Отчет по лифтам
            else if (type == 43)
            {
                DataTable obj = ora.SelectLiftInfo();
                var wb = new XLWorkbook(@"C:\temp\Копия Книга1.xlsx");
                string address = "";
                string municipality = "";
                int rowMove = 1;
                for (int i = 0; i < obj.Rows.Count; i++)
                {
                    rowMove++;
                    if (municipality != obj.Rows[i][0].ToString())
                    {
                        municipality = obj.Rows[i][0].ToString();
                        address = obj.Rows[i][1].ToString();
                        wb.Worksheet(1).Row(rowMove).Cell(1).Value = rowMove - 1;
                        wb.Worksheet(1).Row(rowMove).Cell(2).Value = municipality;
                        wb.Worksheet(1).Row(rowMove).Cell(3).Value = address;
                        if (obj.Rows[i][2] != null)
                            wb.Worksheet(1).Row(rowMove).Cell(4).Value = obj.Rows[i][2].ToString();
                        else
                            wb.Worksheet(1).Row(rowMove).Cell(4).Value = "-";
                        if (obj.Rows[i][3] != null)
                            wb.Worksheet(1).Row(rowMove).Cell(5).Value = obj.Rows[i][3].ToString();
                        else
                            wb.Worksheet(1).Row(rowMove).Cell(5).Value = "-";
                        if (obj.Rows[i][4] != null)
                            wb.Worksheet(1).Row(rowMove).Cell(6).Value = obj.Rows[i][4].ToString();
                        else
                            wb.Worksheet(1).Row(rowMove).Cell(6).Value = "-";
                        if (obj.Rows[i][5] != null)
                            wb.Worksheet(1).Row(rowMove).Cell(7).Value = obj.Rows[i][5].ToString();
                        else
                            wb.Worksheet(1).Row(rowMove).Cell(7).Value = "-";
                        if (obj.Rows[i][6] != null)
                            wb.Worksheet(1).Row(rowMove).Cell(8).Value = obj.Rows[i][6].ToString();
                        else
                            wb.Worksheet(1).Row(rowMove).Cell(8).Value = "-";
                        if (obj.Rows[i][7] != null)
                            wb.Worksheet(1).Row(rowMove).Cell(9).Value = obj.Rows[i][7].ToString();
                        else
                            wb.Worksheet(1).Row(rowMove).Cell(9).Value = "-";
                        if (obj.Rows[i][8] != null)
                            wb.Worksheet(1).Row(rowMove).Cell(10).Value = obj.Rows[i][8].ToString();
                        else
                            wb.Worksheet(1).Row(rowMove).Cell(10).Value = "-";
                        if (obj.Rows[i][9] != null)
                            wb.Worksheet(1).Row(rowMove).Cell(11).Value = obj.Rows[i][9].ToString();
                        else
                            wb.Worksheet(1).Row(rowMove).Cell(11).Value = "-";
                        if (obj.Rows[i][10] != null)
                            wb.Worksheet(1).Row(rowMove).Cell(12).Value = obj.Rows[i][10].ToString();
                        else
                            wb.Worksheet(1).Row(rowMove).Cell(12).Value = "-";
                        if (obj.Rows[i][11] != null)
                            wb.Worksheet(1).Row(rowMove).Cell(13).Value = obj.Rows[i][11].ToString();
                        else
                            wb.Worksheet(1).Row(rowMove).Cell(13).Value = "-";
                    }
                    else
                    {
                        if (address != obj.Rows[i][1].ToString())
                        {
                            address = obj.Rows[i][1].ToString();
                            wb.Worksheet(1).Row(rowMove).Cell(1).Value = rowMove - 1;
                            wb.Worksheet(1).Row(rowMove).Cell(3).Value = address;
                            if (obj.Rows[i][2] != null)
                                wb.Worksheet(1).Row(rowMove).Cell(4).Value = obj.Rows[i][2].ToString();
                            else
                                wb.Worksheet(1).Row(rowMove).Cell(4).Value = "-";
                            if (obj.Rows[i][3] != null)
                                wb.Worksheet(1).Row(rowMove).Cell(5).Value = obj.Rows[i][3].ToString();
                            else
                                wb.Worksheet(1).Row(rowMove).Cell(5).Value = "-";
                            if (obj.Rows[i][4] != null)
                                wb.Worksheet(1).Row(rowMove).Cell(6).Value = obj.Rows[i][4].ToString();
                            else
                                wb.Worksheet(1).Row(rowMove).Cell(6).Value = "-";
                            if (obj.Rows[i][5] != null)
                                wb.Worksheet(1).Row(rowMove).Cell(7).Value = obj.Rows[i][5].ToString();
                            else
                                wb.Worksheet(1).Row(rowMove).Cell(7).Value = "-";
                            if (obj.Rows[i][6] != null)
                                wb.Worksheet(1).Row(rowMove).Cell(8).Value = obj.Rows[i][6].ToString();
                            else
                                wb.Worksheet(1).Row(rowMove).Cell(8).Value = "-";
                            if (obj.Rows[i][7] != null)
                                wb.Worksheet(1).Row(rowMove).Cell(9).Value = obj.Rows[i][7].ToString();
                            else
                                wb.Worksheet(1).Row(rowMove).Cell(9).Value = "-";
                            if (obj.Rows[i][8] != null)
                                wb.Worksheet(1).Row(rowMove).Cell(10).Value = obj.Rows[i][8].ToString();
                            else
                                wb.Worksheet(1).Row(rowMove).Cell(10).Value = "-";
                            if (obj.Rows[i][9] != null)
                                wb.Worksheet(1).Row(rowMove).Cell(11).Value = obj.Rows[i][9].ToString();
                            else
                                wb.Worksheet(1).Row(rowMove).Cell(11).Value = "-";
                            if (obj.Rows[i][10] != null)
                                wb.Worksheet(1).Row(rowMove).Cell(12).Value = obj.Rows[i][10].ToString();
                            else
                                wb.Worksheet(1).Row(rowMove).Cell(12).Value = "-";
                            if (obj.Rows[i][11] != null)
                                wb.Worksheet(1).Row(rowMove).Cell(13).Value = obj.Rows[i][11].ToString();
                            else
                                wb.Worksheet(1).Row(rowMove).Cell(13).Value = "-";
                        }
                        else
                        {
                            wb.Worksheet(1).Row(rowMove).Cell(1).Value = rowMove - 1;
                            if (obj.Rows[i][2] != null)
                                wb.Worksheet(1).Row(rowMove).Cell(4).Value = obj.Rows[i][2].ToString();
                            else
                                wb.Worksheet(1).Row(rowMove).Cell(4).Value = "-";
                            if (obj.Rows[i][3] != null)
                                wb.Worksheet(1).Row(rowMove).Cell(5).Value = obj.Rows[i][3].ToString();
                            else
                                wb.Worksheet(1).Row(rowMove).Cell(5).Value = "-";
                            if (obj.Rows[i][4] != null)
                                wb.Worksheet(1).Row(rowMove).Cell(6).Value = obj.Rows[i][4].ToString();
                            else
                                wb.Worksheet(1).Row(rowMove).Cell(6).Value = "-";
                            if (obj.Rows[i][5] != null)
                                wb.Worksheet(1).Row(rowMove).Cell(7).Value = obj.Rows[i][5].ToString();
                            else
                                wb.Worksheet(1).Row(rowMove).Cell(7).Value = "-";
                            if (obj.Rows[i][6] != null)
                                wb.Worksheet(1).Row(rowMove).Cell(8).Value = obj.Rows[i][6].ToString();
                            else
                                wb.Worksheet(1).Row(rowMove).Cell(8).Value = "-";
                            if (obj.Rows[i][7] != null)
                                wb.Worksheet(1).Row(rowMove).Cell(9).Value = obj.Rows[i][7].ToString();
                            else
                                wb.Worksheet(1).Row(rowMove).Cell(9).Value = "-";
                            if (obj.Rows[i][8] != null)
                                wb.Worksheet(1).Row(rowMove).Cell(10).Value = obj.Rows[i][8].ToString();
                            else
                                wb.Worksheet(1).Row(rowMove).Cell(10).Value = "-";
                            if (obj.Rows[i][9] != null)
                                wb.Worksheet(1).Row(rowMove).Cell(11).Value = obj.Rows[i][9].ToString();
                            else
                                wb.Worksheet(1).Row(rowMove).Cell(11).Value = "-";
                            if (obj.Rows[i][10] != null)
                                wb.Worksheet(1).Row(rowMove).Cell(12).Value = obj.Rows[i][10].ToString();
                            else
                                wb.Worksheet(1).Row(rowMove).Cell(12).Value = "-";
                            if (obj.Rows[i][11] != null)
                                wb.Worksheet(1).Row(rowMove).Cell(13).Value = obj.Rows[i][11].ToString();
                            else
                                wb.Worksheet(1).Row(rowMove).Cell(13).Value = "-";

                        }
                    }
                }
                wb.Save();
            }
            #endregion

            #region 44 Отчет по проценту заполнения домов
            else if (type == 44)
            {
                DataTable obj = ora.SelectPctInfo();
                var wb = new XLWorkbook();
                wb.AddWorksheet("организации");
                string manorg = obj.Rows[0][0].ToString();
                string phone = obj.Rows[0][2].ToString();
                string jurAdres = obj.Rows[0][3].ToString();
                string factAdres = obj.Rows[0][4].ToString();
                int count20 = 0;
                int count40 = 0;
                int count60 = 0;
                int count80 = 0;
                int count100 = 0;
                int rowMove = 1;
                for (int i = 0; i < obj.Rows.Count; i++)
                {
                    if (manorg != obj.Rows[i][0].ToString())
                    {
                        rowMove++;
                        wb.Worksheet(1).Row(rowMove).Cell(1).Value = manorg;
                        wb.Worksheet(1).Row(rowMove).Cell(4).Value = phone;
                        wb.Worksheet(1).Row(rowMove).Cell(5).Value = jurAdres;
                        wb.Worksheet(1).Row(rowMove).Cell(6).Value = factAdres;
                        manorg = obj.Rows[i][0].ToString();
                        phone = obj.Rows[i][2].ToString();
                        jurAdres = obj.Rows[i][3].ToString();
                        factAdres = obj.Rows[i][4].ToString();
                        wb.Worksheet(1).Row(rowMove).Cell(2).Value = "0-20";
                        wb.Worksheet(1).Row(rowMove).Cell(3).Value = count20;
                        count20 = 0;
                        rowMove++;
                        wb.Worksheet(1).Row(rowMove).Cell(2).Value = "21-40";
                        wb.Worksheet(1).Row(rowMove).Cell(3).Value = count40;
                        count40 = 0;
                        rowMove++;
                        wb.Worksheet(1).Row(rowMove).Cell(2).Value = "41-60";
                        wb.Worksheet(1).Row(rowMove).Cell(3).Value = count60;
                        count60 = 0;
                        rowMove++;
                        wb.Worksheet(1).Row(rowMove).Cell(2).Value = "61-80";
                        wb.Worksheet(1).Row(rowMove).Cell(3).Value = count80;
                        count80 = 0;
                        rowMove++;
                        wb.Worksheet(1).Row(rowMove).Cell(2).Value = "81-100";
                        wb.Worksheet(1).Row(rowMove).Cell(3).Value = count100;
                        count100 = 0;
                    }
                    if (Convert.ToDecimal(obj.Rows[i][1].ToString()) <= 20)
                        count20++;
                    else if (Convert.ToDecimal(obj.Rows[i][1].ToString()) <= 40 && Convert.ToDecimal(obj.Rows[i][1].ToString()) > 20)
                        count40++;
                    else if (Convert.ToDecimal(obj.Rows[i][1].ToString()) <= 60 && Convert.ToDecimal(obj.Rows[i][1].ToString()) > 40)
                        count60++;
                    else if (Convert.ToDecimal(obj.Rows[i][1].ToString()) <= 80 && Convert.ToDecimal(obj.Rows[i][1].ToString()) > 60)
                        count80++;
                    else if (Convert.ToDecimal(obj.Rows[i][1].ToString()) > 80)
                        count100++;
                }
                rowMove++;
                wb.Worksheet(1).Row(rowMove).Cell(1).Value = manorg;
                wb.Worksheet(1).Row(rowMove).Cell(4).Value = phone;
                wb.Worksheet(1).Row(rowMove).Cell(5).Value = jurAdres;
                wb.Worksheet(1).Row(rowMove).Cell(6).Value = factAdres;
                wb.Worksheet(1).Row(rowMove).Cell(2).Value = "0-20";
                wb.Worksheet(1).Row(rowMove).Cell(3).Value = count20;
                rowMove++;
                wb.Worksheet(1).Row(rowMove).Cell(2).Value = "21-40";
                wb.Worksheet(1).Row(rowMove).Cell(3).Value = count40;
                rowMove++;
                wb.Worksheet(1).Row(rowMove).Cell(2).Value = "41-60";
                wb.Worksheet(1).Row(rowMove).Cell(3).Value = count60;
                rowMove++;
                wb.Worksheet(1).Row(rowMove).Cell(2).Value = "61-80";
                wb.Worksheet(1).Row(rowMove).Cell(3).Value = count80;
                rowMove++;
                wb.Worksheet(1).Row(rowMove).Cell(2).Value = "81-100";
                wb.Worksheet(1).Row(rowMove).Cell(3).Value = count100;
                wb.SaveAs(@"C:\temp\report"+DateTime.Now.ToShortDateString()+".xlsx");
            }
            #endregion

            #region 45
            else if (type == 45)
            {

                string str = ora.UpdateHouseOwner("6316117366", "17719356");
                if (str != "ЗАГРУЖЕНО")
                    Console.WriteLine(str);
            }
            #endregion

            #region 46
            else if (type == 46)
            {
                string str = ora.InsertHouseManOrg("01.01.2015", "956922");
                if (str != "ЗАГРУЖЕНО")
                    Console.WriteLine(str);
            }
            #endregion

            #region 47
            else if (type == 47)
            {
                string s = "  hello world   !";
                char[] chars = s.ToCharArray();
                int count = 0;
                foreach (char c in chars)
                {
                    if (c == ' ') count++;
                }
                Console.WriteLine(count.ToString());
            }
            #endregion

            #region 48
            else if (type == 48)
            {
                var wb = new XLWorkbook();
                wb.AddWorksheet("организации");
                string str = @"C:\Users\WCSMR-HP\Desktop\МЗСО СО(19-35-15).xml";
                int row = 1;
                using (XmlReader reader = XmlReader.Create(str))
                {
                    while (reader.Read())
                    {
                        string tmp = reader.Name;
                        if (tmp == "Код")
                        {
                            row++;
                            reader.ReadStartElement("Код");
                            wb.Worksheet(1).Row(row).Cell(1).Value = reader.ReadString();
                            reader.ReadEndElement();
                        }
                        if (tmp == "Наименование")
                        {
                            reader.ReadStartElement("Наименование");
                            wb.Worksheet(1).Row(row).Cell(2).Value = reader.ReadString();
                            reader.ReadEndElement();
                        }
                        if (tmp == "ИНН")
                        {
                            try
                            {
                                reader.ReadStartElement("ИНН");
                                wb.Worksheet(1).Row(row).Cell(3).Value = reader.ReadString();
                                reader.ReadEndElement();
                            }
                            catch { }
                        }
                    }
                }
                wb.SaveAs(@"C:\temp\temp22222.xlsx");
                    
            }
            #endregion

            #region 49
            else if (type == 49)
            {
                DataTable dtHouse = new DataTable();
                DataTable dtPeople = new DataTable();
                var wb2 = new XLWorkbook(@"C:\temp\Копия Реестр квартир и собсвтенников_А9_10 домов домов.xlsx");
                for (int i = 476; i <= 1036; i++)
                {
                    if(i%100 == 0)
                        Console.WriteLine(i.ToString());
                    if (wb2.Worksheet(1).Row(i).Cell(2).Value != null && wb2.Worksheet(1).Row(i).Cell(2).Value.ToString() != "")
                    {
                        try
                        {
                            string str = ora.InsertPeople3(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Replace('(', ' ').Replace(')', ' ').Trim(),
                                Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim(),
                                Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim(),
                                Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value));
                            if (str != "ЗАГРУЖЕНО")
                                Console.WriteLine(str);
                        }
                        catch
                        {

                        }
                    }

                }
                wb2.Save();
            }
            #endregion

            #region 50
            else if (type == 50)
            {
                var wb2 = new XLWorkbook(@"C:\temp\сведения о квартирах(1).xlsx");
                for (int i = 2; i <= 55; i++)
                {
                    try
                    {
                        string str = ora.InsertPeople4("8800494",
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim(),
                            "",
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim(),
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim(),
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim()
                            );
                        if (str != "ЗАГРУЖЕНО")
                            Console.WriteLine(str);
                    }
                    catch
                    {

                    }
                }
                wb2.Save();
            }
            #endregion

            //Перенос тарифов из Информикса в Постгре
            #region 51
            else if (type == 51)
            {
                var wb2 = new XLWorkbook(@"C:\temp\tarif.xlsx");//172-178
                int start = 0;
                int end = 0;
                Console.Write("введите наименование БД:");
                string database = Console.ReadLine();
                Console.Write("введите наименование банка:");
                string bank = Console.ReadLine();
                Console.Write("Введите начальную строку:");
                start = Convert.ToInt32(Console.ReadLine());
                Console.Write("Введите конечную строку:");
                end = Convert.ToInt32(Console.ReadLine());
                //список банков
                List<string> prefs = new List<string>();
                prefs.Add(bank);
                for (int i = start; i <= end; i++)
                {
                    pg.InsertTarif(database,
                        Convert.ToInt32(wb2.Worksheet(1).Row(i).Cell(2).Value),
                        Convert.ToInt32(wb2.Worksheet(1).Row(i).Cell(5).Value),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(7).Value),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(8).Value),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(9).Value),
                        prefs,
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(10).Value),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(11).Value),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(12).Value),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(13).Value),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(14).Value));
                }
            }
            #endregion

            //Перенос нормативов из Информикса в Постгре
            #region 52
            else if (type == 52)
            {
                //список банков
                List<string> prefs = new List<string>();
                prefs.Add("bill01");
                DataRow row;
                //создаем таблицу prm_name 
                var wbPrmName = new XLWorkbook(@"C:\temp\нормативы\prm_name.xlsx");
                DataTable dtPrmName = new System.Data.DataTable();
                dtPrmName.Columns.Add("1");
                dtPrmName.Columns.Add("2");
                for (int i = 2; i <= 12; i++)
                {
                    if (Convert.ToString(wbPrmName.Worksheet(1).Row(i).Cell(2).Value) != "")
                    {
                        row = dtPrmName.NewRow();
                        row["1"] = Convert.ToString(wbPrmName.Worksheet(1).Row(i).Cell(1).Value);
                        row["2"] = Convert.ToString(wbPrmName.Worksheet(1).Row(i).Cell(2).Value);
                        dtPrmName.Rows.Add(row);
                    }
                }

                //создаем таблицу prm_13
                var wbPrm13 = new XLWorkbook(@"C:\temp\нормативы\prm_13.xlsx");
                DataTable dtPrm13 = new System.Data.DataTable();
                dtPrm13.Columns.Add("1");
                dtPrm13.Columns.Add("2");
                dtPrm13.Columns.Add("3");
                dtPrm13.Columns.Add("4");
                dtPrm13.Columns.Add("5");
                dtPrm13.Columns.Add("6");
                dtPrm13.Columns.Add("7");
                for (int i = 2; i <= 51; i++)
                {
                    if (Convert.ToString(wbPrm13.Worksheet(1).Row(i).Cell(2).Value) != "")
                    {
                        row = dtPrm13.NewRow();
                        row["1"] = Convert.ToString(wbPrm13.Worksheet(1).Row(i).Cell(3).Value);
                        row["2"] = Convert.ToString(wbPrm13.Worksheet(1).Row(i).Cell(4).Value);
                        row["3"] = Convert.ToString(wbPrm13.Worksheet(1).Row(i).Cell(5).Value);
                        row["4"] = Convert.ToString(wbPrm13.Worksheet(1).Row(i).Cell(6).Value);
                        row["5"] = Convert.ToString(wbPrm13.Worksheet(1).Row(i).Cell(7).Value);
                        row["6"] = Convert.ToString(wbPrm13.Worksheet(1).Row(i).Cell(8).Value);
                        row["7"] = Convert.ToString(wbPrm13.Worksheet(1).Row(i).Cell(8).Value);
                        dtPrm13.Rows.Add(row);
                    }
                }

                //создаем таблицу rex_value 
                var wbResValue = new XLWorkbook(@"C:\temp\нормативы\res_value.xlsx");
                DataTable dtResValue = new System.Data.DataTable();
                dtResValue.Columns.Add("1");
                dtResValue.Columns.Add("2");
                dtResValue.Columns.Add("3");
                dtResValue.Columns.Add("4");
                for (int i = 2; i <= 2679; i++)
                {
                    if (Convert.ToString(wbResValue.Worksheet(1).Row(i).Cell(2).Value) != "")
                    {
                        row = dtResValue.NewRow();
                        row["1"] = Convert.ToString(wbResValue.Worksheet(1).Row(i).Cell(1).Value);
                        row["2"] = Convert.ToString(wbResValue.Worksheet(1).Row(i).Cell(2).Value);
                        row["3"] = Convert.ToString(wbResValue.Worksheet(1).Row(i).Cell(3).Value);
                        row["4"] = Convert.ToString(wbResValue.Worksheet(1).Row(i).Cell(4).Value);
                        dtResValue.Rows.Add(row);
                    }
                }

                //создаем таблицу res_x
                var wbResX = new XLWorkbook(@"C:\temp\нормативы\res_x.xlsx");
                DataTable dtResX = new System.Data.DataTable();
                dtResX.Columns.Add("1");
                dtResX.Columns.Add("2");
                dtResX.Columns.Add("3");
                for (int i = 2; i <= 722; i++)
                {
                    if (Convert.ToString(wbResX.Worksheet(1).Row(i).Cell(2).Value) != "")
                    {
                        row = dtResX.NewRow();
                        row["1"] = Convert.ToString(wbResX.Worksheet(1).Row(i).Cell(1).Value);
                        row["2"] = Convert.ToString(wbResX.Worksheet(1).Row(i).Cell(2).Value);
                        row["3"] = Convert.ToString(wbResX.Worksheet(1).Row(i).Cell(3).Value);
                        dtResX.Rows.Add(row);
                    }
                }

                //создаем таблицу res_y
                var wbResY = new XLWorkbook(@"C:\temp\нормативы\res_y.xlsx");
                DataTable dtResY = new System.Data.DataTable();
                dtResY.Columns.Add("1");
                dtResY.Columns.Add("2");
                dtResY.Columns.Add("3");
                for (int i = 2; i <= 906; i++)
                {
                    if (Convert.ToString(wbResY.Worksheet(1).Row(i).Cell(2).Value) != "")
                    {
                        row = dtResY.NewRow();
                        row["1"] = Convert.ToString(wbResY.Worksheet(1).Row(i).Cell(1).Value);
                        row["2"] = Convert.ToString(wbResY.Worksheet(1).Row(i).Cell(2).Value);
                        row["3"] = Convert.ToString(wbResY.Worksheet(1).Row(i).Cell(3).Value);
                        dtResY.Rows.Add(row);
                    }
                }

                //создаем таблицу resolution
                var wbResolution = new XLWorkbook(@"C:\temp\нормативы\resolution.xlsx");
                DataTable dtResolution = new System.Data.DataTable();
                dtResolution.Columns.Add("1");
                dtResolution.Columns.Add("2");
                dtResolution.Columns.Add("3");
                for (int i = 2; i <= 98; i++)
                {
                    if (Convert.ToString(wbResolution.Worksheet(1).Row(i).Cell(2).Value) != "")
                    {
                        row = dtResolution.NewRow();
                        row["1"] = Convert.ToString(wbResolution.Worksheet(1).Row(i).Cell(1).Value);
                        row["2"] = Convert.ToString(wbResolution.Worksheet(1).Row(i).Cell(2).Value);
                        row["3"] = Convert.ToString(wbResolution.Worksheet(1).Row(i).Cell(3).Value);
                        dtResolution.Rows.Add(row);
                    }
                }
                Dictionary<int, int> prmNameConvert = new Dictionary<int, int>();
                for (int i = 0; i < dtPrmName.Rows.Count; i++)
                {
                    int prmNameID = pg.InsertPrmName(dtPrmName.Rows[i][1].ToString(), prefs);
                    prmNameConvert.Add(Convert.ToInt32(dtPrmName.Rows[i][0]), prmNameID);
                }
                Dictionary<int, int> resolutionConvert = new Dictionary<int, int>();
                for (int i = 0; i < dtResolution.Rows.Count; i++)
                {
                    int resolutionID = pg.InsertResolution(dtResolution.Rows[i][1].ToString(),
                                                            dtResolution.Rows[i][2].ToString(),
                                                            prefs);
                    int id = Convert.ToInt32(dtResolution.Rows[i][0]);
                    resolutionConvert.Add(id, resolutionID);
                    for (int j = 0; j < dtResX.Rows.Count; j++)
                    {
                        if (Convert.ToInt32(dtResX.Rows[j][0]) == Convert.ToInt32(dtResolution.Rows[i][0]))
                        {
                            pg.InsertResX(resolutionID, Convert.ToInt32(dtResX.Rows[j][1]), dtResX.Rows[j][2].ToString(), prefs);
                        }
                    }
                    for (int j = 0; j < dtResY.Rows.Count; j++)
                    {
                        if (Convert.ToInt32(dtResY.Rows[j][0]) == Convert.ToInt32(dtResolution.Rows[i][0]))
                        {
                            pg.InsertResY(resolutionID, Convert.ToInt32(dtResY.Rows[j][1]), dtResY.Rows[j][2].ToString(), prefs);
                        }
                    }
                    for (int j = 0; j < dtResValue.Rows.Count; j++)
                    {
                        if (Convert.ToInt32(dtResValue.Rows[j][0]) == Convert.ToInt32(dtResolution.Rows[i][0]))
                        {
                            pg.InsertResValue(resolutionID, Convert.ToInt32(dtResValue.Rows[j][1]), Convert.ToInt32(dtResValue.Rows[j][2]), dtResValue.Rows[j][2].ToString(), prefs);
                        }
                    }
                }
                foreach (KeyValuePair<int, int> nzp_prm in prmNameConvert)
                {
                    foreach (KeyValuePair<int, int> val_prm in resolutionConvert)
                    {
                        for (int i = 0; i < dtPrm13.Rows.Count; i++)
                        {
                            if (Convert.ToInt32(dtPrm13.Rows[i][0]) == nzp_prm.Key && Convert.ToInt32(dtPrm13.Rows[i][3]) == val_prm.Key)
                            {
                                pg.InsertPrm13(nzp_prm.Value, dtPrm13.Rows[i][1].ToString(), dtPrm13.Rows[i][2].ToString(), val_prm.Value, dtPrm13.Rows[i][4].ToString(), dtPrm13.Rows[i][6].ToString(), prefs);
                            }
                        }
                    }
                }
                wbPrm13.Save();
                wbPrmName.Save();
                wbResolution.Save();
                wbResValue.Save();
                wbResX.Save();
                wbResY.Save();
            }
            #endregion

            //Перенос нормативов из Информикса в Постгре
            #region 53
            else if (type == 53)
            {
                var wbPrmValues = new XLWorkbook(@"C:\Work\billPGDataOld\bill01_res_values.xlsx");
                for (int i = 1; i <= 3376; i++)
                {
                    pg.InsertResValues2(Convert.ToString(wbPrmValues.Worksheet(1).Row(i).Cell(1).Value),
                        Convert.ToString(wbPrmValues.Worksheet(1).Row(i).Cell(2).Value),
                        Convert.ToString(wbPrmValues.Worksheet(1).Row(i).Cell(3).Value),
                        Convert.ToString(wbPrmValues.Worksheet(1).Row(i).Cell(4).Value).Trim(),
                        Convert.ToString(wbPrmValues.Worksheet(1).Row(i).Cell(5).Value), "bill01");
                }

                var wbPrmValues2 = new XLWorkbook(@"C:\Work\billPGDataOld\fbill_res_values.xlsx");
                for (int i = 1; i <= 3216; i++)
                {
                    pg.InsertResValues2(Convert.ToString(wbPrmValues2.Worksheet(1).Row(i).Cell(1).Value),
                        Convert.ToString(wbPrmValues2.Worksheet(1).Row(i).Cell(2).Value),
                        Convert.ToString(wbPrmValues2.Worksheet(1).Row(i).Cell(3).Value),
                        Convert.ToString(wbPrmValues2.Worksheet(1).Row(i).Cell(4).Value),
                        Convert.ToString(wbPrmValues2.Worksheet(1).Row(i).Cell(5).Value), "fbill");
                }

                var wbResX = new XLWorkbook(@"C:\Work\billPGDataOld\bill01_res_x.xlsx");
                for (int i = 1; i <= 330; i++)
                {
                    pg.InsertResX2(Convert.ToString(wbResX.Worksheet(1).Row(i).Cell(1).Value),
                        Convert.ToString(wbResX.Worksheet(1).Row(i).Cell(2).Value),
                        Convert.ToString(wbResX.Worksheet(1).Row(i).Cell(3).Value),
                        Convert.ToString(wbResX.Worksheet(1).Row(i).Cell(4).Value), "bill01");
                }

                var wbResX2 = new XLWorkbook(@"C:\Work\billPGDataOld\fbill_res_x.xlsx");
                for (int i = 1; i <= 326; i++)
                {
                    pg.InsertResX2(Convert.ToString(wbResX2.Worksheet(1).Row(i).Cell(1).Value),
                        Convert.ToString(wbResX2.Worksheet(1).Row(i).Cell(2).Value),
                        Convert.ToString(wbResX2.Worksheet(1).Row(i).Cell(3).Value),
                        Convert.ToString(wbResX2.Worksheet(1).Row(i).Cell(4).Value), "fbill");
                }

                var wbResY = new XLWorkbook(@"C:\Work\billPGDataOld\bill01_res_y.xlsx");
                for (int i = 1; i <= 992; i++)
                {
                    pg.InsertResY2(Convert.ToString(wbResY.Worksheet(1).Row(i).Cell(1).Value),
                        Convert.ToString(wbResY.Worksheet(1).Row(i).Cell(2).Value),
                        Convert.ToString(wbResY.Worksheet(1).Row(i).Cell(3).Value),
                        Convert.ToString(wbResY.Worksheet(1).Row(i).Cell(4).Value), "bill01");
                }

                var wbResY2 = new XLWorkbook(@"C:\Work\billPGDataOld\fbill_res_y.xlsx");
                for (int i = 1; i <= 959; i++)
                {
                    pg.InsertResY2(Convert.ToString(wbResY2.Worksheet(1).Row(i).Cell(1).Value),
                        Convert.ToString(wbResY2.Worksheet(1).Row(i).Cell(2).Value),
                        Convert.ToString(wbResY2.Worksheet(1).Row(i).Cell(3).Value),
                        Convert.ToString(wbResY2.Worksheet(1).Row(i).Cell(4).Value), "fbill");
                }

                var wbResolution = new XLWorkbook(@"C:\Work\billPGDataOld\bill01_resolution.xlsx");
                for (int i = 1; i <= 115; i++)
                {
                    pg.InsertResolution2(Convert.ToString(wbResolution.Worksheet(1).Row(i).Cell(1).Value),
                        Convert.ToString(wbResolution.Worksheet(1).Row(i).Cell(2).Value),
                        Convert.ToString(wbResolution.Worksheet(1).Row(i).Cell(3).Value), "bill01");
                }

                var wbResolution2 = new XLWorkbook(@"C:\Work\billPGDataOld\fbill_resolution.xlsx");
                for (int i = 1; i <= 118; i++)
                {
                    pg.InsertResolution2(Convert.ToString(wbResolution2.Worksheet(1).Row(i).Cell(1).Value),
                        Convert.ToString(wbResolution2.Worksheet(1).Row(i).Cell(2).Value),
                        Convert.ToString(wbResolution2.Worksheet(1).Row(i).Cell(3).Value), "fbill");
                }

                var wbPrmName = new XLWorkbook(@"C:\Work\billPGDataOld\fbill_prm_name.xlsx");
                for (int i = 1; i <= 18; i++)
                {
                    pg.InsertPrmName2(Convert.ToString(wbPrmName.Worksheet(1).Row(i).Cell(1).Value),
                        Convert.ToString(wbPrmName.Worksheet(1).Row(i).Cell(2).Value), "fbill");
                }
                for (int i = 1; i <= 18; i++)
                {
                    pg.InsertPrmName2(Convert.ToString(wbPrmName.Worksheet(1).Row(i).Cell(1).Value),
                        Convert.ToString(wbPrmName.Worksheet(1).Row(i).Cell(2).Value), "bill01");
                }
            }
            #endregion

            //ЭЖКХ. Перенос собственников по одному дому со SPLIT-ом по квартире
            #region 54
            else if (type == 54)
            {
                DataTable dtHouse = new DataTable();
                DataTable dtPeople = new DataTable();
                var wb2 = new XLWorkbook(@"C:\temp\Копия общие сведения.xlsx");
                string[] stringSeparators = new string[] { ", кв." };
                for (int i = 2; i <= 995; i++)
                {
                    try
                    {
                        //string str = ora.InsertPeople2("9700504",
                        //    Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim().Split(stringSeparators, StringSplitOptions.None)[1],
                        //    Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value).Trim(), 
                        //    Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value), 
                        //    Convert.ToString(wb2.Worksheet(1).Row(i).Cell(7).Value),
                        //    Convert.ToString(wb2.Worksheet(1).Row(i).Cell(8).Value),
                        //    Convert.ToString(wb2.Worksheet(1).Row(i).Cell(12).Value).Trim());
                        //if (str != "ЗАГРУЖЕНО")
                        //    Console.WriteLine(str);
                    }
                    catch
                    {
                       
                    }

                }
                wb2.Save();
            }
            #endregion

            //Поставщики коммунальных услуг с добавление дома
            #region 55
            else if (type == 55)
            {
                var wb2 = new XLWorkbook(@"C:\temp\Копия Копия Сопоставление.xlsx");
                for (int i = 2; i <= 1726; i++)
                {
                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim() != "" && Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim() != "ИЖС")
                    {
                        string str = ora.InsertCommunalOrg2("18517527", Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim());
                        if (str != "ЗАГРУЖЕНО")
                            Console.WriteLine(str);
                    }
                }
            }
            #endregion

            //Загрузка платежей в стройку
            #region 56
            else if (type == 56)
            {
                var wb2 = new XLWorkbook(@"C:\temp\ПП для загрузки.xlsx");
                for (int i = 4; i <= 38; i++)
                {
                    string str = pg.InsertPayment(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim().Replace(',','.'),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(9).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(10).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(11).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(12).Value).Trim());
                    if (str != "ЗАГРУЖЕНО")
                        Console.WriteLine(str);
                    else
                        Console.WriteLine("ЗАГРУЖЕНО = " + Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim());
                }
            }
            #endregion

            //Перенос жильцов из одного дома в другой
            #region 57
            else if (type == 57)
            {
                string str = ora.RemovePeople("339217", "328756");
                if (str != "ЗАГРУЖЕНО")
                    Console.WriteLine(str);
            }
            #endregion

            //Загрузка домов и жильцов в биллинг из Excel
            #region 58
            else if (type == 58)
            {
                var wb2 = new XLWorkbook(@"C:\temp\!БАЗА для загрузки.xlsx");
                string[] stringSeparators = new string[] { "в." };
                string ul = "";
                int nzp_ul = 0;
                string dom = "";
                int nzp_dom = 0;
                for (int i = 2; i <= 822; i++)
                {
                    if (ul != Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim())
                    {
                        ul = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim();
                        string temp = pg.SelectNzpUl(ul);
                        if (temp.Split('|')[0] == "0")
                        {
                            Console.WriteLine(ul + ": " + temp.Split('|')[1]);
                        }
                        else
                        {
                            nzp_ul = Convert.ToInt32(temp.Split('|')[0]);
                            Console.WriteLine(ul + ": " + nzp_ul);
                            dom = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim();
                            nzp_dom = pg.InsertDom(nzp_ul, dom);
                            Console.WriteLine(dom + ": " + nzp_dom);
                        }
                    }
                    if (dom != Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim())
                    {
                        dom = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim();
                        nzp_dom = pg.InsertDom(nzp_ul, dom);
                        Console.WriteLine(dom + ": " + nzp_dom);
                    }
                    string nkvar = "";
                    if(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim().Substring(0,3) == "Кв.")
                        nkvar = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim().Split(stringSeparators, StringSplitOptions.None)[1].Trim();
                    else
                        nkvar =Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim();
                    int num_ls = 0;
                    int number;
                    int subLs = 0;
                    string num_ls_full = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim().Split('№')[1].Trim();
                    for (int j = 0; j < num_ls_full.Length; j++)
                    {
                        if (Int32.TryParse(num_ls_full.Substring(j, 1), out number))
                            subLs++;
                        else
                            break;
                    }
                    if(subLs > 0)
                        num_ls = Convert.ToInt32(num_ls_full.Substring(0, subLs));

                    int ikvar = 0;
                    int subKvar = 0;
                    for (int j = 0; j < nkvar.Length; j++)
                    {
                        if (Int32.TryParse(nkvar.Substring(j, 1), out number))
                            subKvar++;
                        else
                            break;
                    }
                    if (subKvar > 0)
                        ikvar = Convert.ToInt32(nkvar.Substring(0, subKvar));

                    int nzp_kvar = pg.InsertKvar(nzp_dom, nkvar, num_ls, Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value).Trim(), ikvar);
                    pg.InsertDateOpen(nzp_kvar, Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value).Trim());
                    if (wb2.Worksheet(1).Row(i).Cell(7).Value != null && Convert.ToString(wb2.Worksheet(1).Row(i).Cell(7).Value).Trim() != "")
                        pg.InsertPrm1(nzp_kvar, Convert.ToString(wb2.Worksheet(1).Row(i).Cell(7).Value).Trim(), 4);
                    if (wb2.Worksheet(1).Row(i).Cell(8).Value != null && Convert.ToString(wb2.Worksheet(1).Row(i).Cell(8).Value).Trim() != "")
                        pg.InsertPrm1(nzp_kvar, Convert.ToString(wb2.Worksheet(1).Row(i).Cell(8).Value).Trim(), 6);
                    if (wb2.Worksheet(1).Row(i).Cell(9).Value != null && Convert.ToString(wb2.Worksheet(1).Row(i).Cell(9).Value).Trim() != "")
                        pg.InsertPrm1(nzp_kvar, Convert.ToString(wb2.Worksheet(1).Row(i).Cell(9).Value).Trim(), 5);
                }
            }
            #endregion

            //Перенос домов и жильцов из Информикса в Постгри
            #region 59
            else if (type == 59)
            {
                var aukCounters = new XLWorkbook(@"C:\temp\АУК\aukCounters.xlsx");
                var aukCountersSpis = new XLWorkbook(@"C:\temp\АУК\aukCountersSpis.xlsx");
                var aukCountersSpis2 = new XLWorkbook(@"C:\temp\АУК\aukCountersSpis2.xlsx");
                var aukCountTypes = new XLWorkbook(@"C:\temp\АУК\aukCountTypes.xlsx");
                var aukDom = new XLWorkbook(@"C:\temp\АУК\aukDom.xlsx");
                var aukKvar = new XLWorkbook(@"C:\temp\АУК\aukKvar.xlsx");
                var aukLit = new XLWorkbook(@"C:\temp\АУК\aukLit.xlsx");
                var aukPrm1 = new XLWorkbook(@"C:\temp\АУК\aukPrm1.xlsx");
                var aukPrm2 = new XLWorkbook(@"C:\temp\АУК\aukPrm2.xlsx");
                var aukPrm3 = new XLWorkbook(@"C:\temp\АУК\aukPrm3.xlsx");
                for (int i = 2; i <= 5; i++)//24201
                {
                    Console.WriteLine(Convert.ToString(aukCounters.Worksheet(1).Row(i).Cell(6).Value).Trim());
                }
            }
            #endregion

            //Групповой ввод характеристик жилья
            #region 60
            else if (type == 60)
            {
                Start:
                Console.Write("Введите наименование БД:");
                string database = Console.ReadLine();
                Console.Write("Введите наименование параметра:");
                string prm_name = Console.ReadLine();
                string curFile = @"C:\temp\Exp\"+prm_name+".xlsx";
                int start = 0;
                int end = 0;
                Console.Write("Введите начальную строку:");
                start = Convert.ToInt32(Console.ReadLine());
                Console.Write("Введите конечную строку:");
                end = Convert.ToInt32(Console.ReadLine());
                if (File.Exists(curFile))
                {
                    Console.WriteLine("Фаил найден: " + @"C:\temp\Exp\" + prm_name + ".xlsx");
                    string val_prm = "";
                    var book = new XLWorkbook(@"C:\temp\Exp\" + prm_name + ".xlsx");
                    Dictionary<string, string> prms = pg.SelectPrms(database, prm_name);
                    if (prms != null)
                    {
                        Console.WriteLine("Параметр найден: prm_num = " + prms["prm_num"] + ", nzp_prm = " + prms["nzp_prm"]);
                        for (int i = start; i <= end; i++)
                        {
                            int nzp_dom = 0;
                            if (Convert.ToString(book.Worksheet(1).Row(i).Cell(1).Value).Trim() == "" && Convert.ToString(book.Worksheet(1).Row(i).Cell(2).Value).Trim() == "")
                                break;
                            if (Convert.ToString(book.Worksheet(1).Row(i).Cell(1).Value).Trim() != "" && val_prm != Convert.ToString(book.Worksheet(1).Row(i).Cell(1).Value).Trim())
                                val_prm = Convert.ToString(book.Worksheet(1).Row(i).Cell(1).Value).Trim();
                            string dom = pg.SelectDom(database, Convert.ToString(book.Worksheet(1).Row(i).Cell(2).Value).Trim(), Convert.ToString(book.Worksheet(1).Row(i).Cell(3).Value).Trim());
                            if (dom.Split('|')[0] == "0")
                            {
                                Console.WriteLine(Convert.ToString(book.Worksheet(1).Row(i).Cell(2).Value).Trim() + ". д" +
                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(3).Value).Trim() + ": " + dom.Split('|')[1]);
                                goto Start;
                            }
                            else
                            {
                                
                                nzp_dom = Convert.ToInt32(dom.Split('|')[0]);
                                Console.WriteLine("Дом найден: " + Convert.ToString(book.Worksheet(1).Row(i).Cell(2).Value).Trim() + ". д" +
                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(3).Value).Trim() + " = " + nzp_dom);
                                DataTable kvars = pg.SelectKvar(database, nzp_dom);
                                if(kvars != null)
                                {
                                    Console.WriteLine("Найдено " + kvars.Rows.Count + " квартир(ы)");
                                    for (int j = 0; j < kvars.Rows.Count; j++)
                                    {
                                        pg.UpdateParams(database, prms["nzp_prm"], prms["prm_num"], kvars.Rows[j][0].ToString(), Convert.ToString(book.Worksheet(1).Row(i).Cell(4).Value).Trim(), val_prm);
                                    }
                                }
                                else
                                {
                                    Console.WriteLine("Квартиры не найдены");
                                    goto Start;
                                }
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("Параметр не найден");
                        goto Start;
                    }
                }
                else
                {
                    Console.WriteLine("Фаил не найден");
                    goto Start;
                }
            }
            #endregion

            //из Экселя в текст
            #region 61
            else if (type == 61)
            {
                var book = new XLWorkbook(@"C:\temp\Гастелло 2  номера.xlsx");
                StreamWriter sw1 = new StreamWriter(@"C:\temp\nkvar2.txt", false);
                for (int i = 1; i <= 500; i++)
                {
                    if (Convert.ToString(book.Worksheet(2).Row(i).Cell(1).Value).Trim() == "")
                        break;
                    sw1.Write("'" + Convert.ToString(book.Worksheet(2).Row(i).Cell(1).Value).Trim() + "',");
                }
                sw1.Close();
            }
            #endregion

            //сорттировка листа
            #region 62
            else if (type == 62)
            {
                List<int> str = new List<int>();
                str.Add(1);
                str.Add(2);
                str.Add(3);
                str.Add(4);
                str.Add(5);
                str.Add(6);
                str.Add(7);
                int pos = 0;
                int posAfter = 0;
                List<int> str2 = new List<int>();
                foreach (int bs in str)
                {
                    if (bs== 3)
                    {
                        str2.Insert(pos, bs);
                        posAfter = pos+1;
                    }
                    else if (bs == 7)
                    {
                        str2.Insert(posAfter, bs);
                    }
                    else
                    {
                        str2.Insert(pos, bs);
                    }
                    pos++;
                }
                foreach(int i in str2)
                {
                    Console.WriteLine(i);
                }
            }
            #endregion

            //Перенос тарифов из Информикса в Постгре
            #region 63
            else if (type == 63)
            {
                var wb2 = new XLWorkbook(@"C:\temp\грохнуть ИПУ.xlsx");
                //список банков
                List<string> prefs = new List<string>();
                bool hvs = true;
                bool gvs = true;
                for (int i = 4; i <= 316; i++)
                {
                    if (wb2.Worksheet(1).Row(i).Cell(5).Value != null && Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value) != "") 
                        hvs = true;
                    else
                        hvs = false;
                    if (wb2.Worksheet(1).Row(i).Cell(6).Value != null && Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value) != "")
                        gvs = true;
                    else
                        gvs = false;
                    int rows_count = pg.DelCounter(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value),
                                    Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value),
                                    Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value),
                                    hvs, gvs);
                    if (rows_count == 0)
                    {
                        wb2.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                }
                wb2.Save();
            }
            #endregion

            //XML в Биллинг
            #region 64
            else if (type == 64)
            {
                var book = new XLWorkbook(@"C:\temp\1С.xlsx");
                var VersionXML = new XLWorkbook(@"C:\temp\VersionXML.xlsx");
                Dictionary<string, string> months = new Dictionary<string, string>();
                decimal result = 0;
                months.Add("Январь", "01");
                months.Add("Февраль", "02");
                months.Add("Март", "03");
                months.Add("Апрель", "04");
                months.Add("Май", "05");
                months.Add("Июнь", "06");
                months.Add("Июль", "07");
                months.Add("Август", "08");
                months.Add("Сентябрь", "09");
                months.Add("Октябрь", "10");
                months.Add("Ноябрь", "11");
                months.Add("Декабрь", "12");
                string year = Convert.ToString(book.Worksheet(1).Row(1).Cell(1).Value).Trim().Split(' ')[3];
                string month = months[Convert.ToString(book.Worksheet(1).Row(1).Cell(1).Value).Trim().Split(' ')[2]];
                int version = 0;
                int subVersion = 0;
                for (int i = 1; i <= 1000; i++)
                {
                    if (Convert.ToString(VersionXML.Worksheet(1).Row(i).Cell(1).Value).Trim() == "")
                    {
                        VersionXML.Worksheet(1).Row(i).Cell(1).Value = year;
                        VersionXML.Worksheet(1).Row(i).Cell(2).Value = month;
                        if(subVersion+1 == 10)
                        {
                            version++;
                            subVersion = 1;
                        }
                        else
                        {
                            subVersion++;
                        }
                        VersionXML.Worksheet(1).Row(i).Cell(3).Value = version;
                        VersionXML.Worksheet(1).Row(i).Cell(4).Value = subVersion;
                        break;
                    }
                    if (Convert.ToString(VersionXML.Worksheet(1).Row(i).Cell(1).Value).Trim() == year &&
                        Convert.ToString(VersionXML.Worksheet(1).Row(i).Cell(2).Value).Trim() == month)
                    {
                        version = Convert.ToInt32(VersionXML.Worksheet(1).Row(i).Cell(3).Value);
                        subVersion = Convert.ToInt32(VersionXML.Worksheet(1).Row(i).Cell(4).Value);
                    }
                }
                XmlTextWriter myXml = new XmlTextWriter(@"C:\temp\Начисление_" + year + "-" + month + ".xml", System.Text.Encoding.Default);
                myXml.Formatting = Formatting.Indented;
                myXml.WriteStartDocument(false);
                myXml.WriteStartElement("ВерсияФорматаФайла");
                myXml.WriteElementString("ВерсияФайла", "2014.12." + version + "." + subVersion);
                myXml.WriteElementString("НаименованиеПО", "Биллинг");
                myXml.WriteElementString("ВерсияПО", "2014-2014");
                myXml.WriteStartElement("УчреждениеОтправитель");
                myXml.WriteElementString("НаименованиеУчреждения", "АУК");
                myXml.WriteElementString("ДатаФормирования", DateTime.Now.ToShortDateString());
                for (int i = 3; i <= 100000; i++)
                {
                    if (Convert.ToString(book.Worksheet(1).Row(i).Cell(1).Value).Trim() == "")
                        break;
                    myXml.WriteStartElement("Начисление");
                    myXml.WriteElementString("НомерПП", Convert.ToString(book.Worksheet(1).Row(i).Cell(1).Value));
                    myXml.WriteElementString("Номенклатура", Convert.ToString(book.Worksheet(1).Row(i).Cell(2).Value));
                    myXml.WriteElementString("Улица", Convert.ToString(book.Worksheet(1).Row(i).Cell(3).Value));
                    myXml.WriteElementString("Дом", Convert.ToString(book.Worksheet(1).Row(i).Cell(4).Value));
                    myXml.WriteElementString("Квартира", Convert.ToString(book.Worksheet(1).Row(i).Cell(5).Value));
                    myXml.WriteElementString("НомерЛС", Convert.ToString(book.Worksheet(1).Row(i).Cell(6).Value));
                    myXml.WriteElementString("Контрагент", Convert.ToString(book.Worksheet(1).Row(i).Cell(7).Value));
                    myXml.WriteElementString("ДоговорКонтрагента", " ");
                    myXml.WriteElementString("Льгота", Convert.ToString(book.Worksheet(1).Row(i).Cell(9).Value));
                    myXml.WriteElementString("ВариантПоставки", Convert.ToString(book.Worksheet(1).Row(i).Cell(10).Value));
                    myXml.WriteElementString("СуммаНачислений", Convert.ToString(book.Worksheet(1).Row(i).Cell(11).Value));
                    result += Convert.ToDecimal(book.Worksheet(1).Row(i).Cell(11).Value);
                    myXml.WriteElementString("СчетУчетаУслуг", Convert.ToString(book.Worksheet(1).Row(i).Cell(12).Value));
                    myXml.WriteElementString("Субконто", Convert.ToString(book.Worksheet(1).Row(i).Cell(13).Value));
                    myXml.WriteEndElement();
                }
                myXml.WriteElementString("Итого", result.ToString());
                myXml.WriteEndElement();
                myXml.WriteEndElement();
                myXml.Flush();
                myXml.Close();
                VersionXML.Save();
                book.Save();
            }
            #endregion

            //согрузка ПС КАРТОТЕКА
            #region 65
            else if (type == 65)
            {
                var book = new XLWorkbook(@"C:\temp\для СОБИТС (ПВС)-1.xlsx");
                for (int i = 34; i <= 151; i++)
                {
                    string nzp_kvar = pg.SelectNzpKvar(Convert.ToString(book.Worksheet(1).Row(i).Cell(1).Value).Trim());

                    if (nzp_kvar.Split('|')[0] == "0")
                    {
                        book.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                        book.Worksheet(1).Row(i).Cell(21).Value = nzp_kvar.Split('|')[1];
                    }
                    else
                    {
                        int nzp_gil = pg.InsertGil();
                        int nzp_rod = 0;
                        #region nzp_rod
                        switch (Convert.ToString(book.Worksheet(1).Row(i).Cell(20).Value).Trim())
                        {
                            case "брат":
                            {
                                nzp_rod = 587;
                                break;
                            }
                            case "внук":
                            {
                                nzp_rod = 560;
                                break;
                            }
                            case "внучка":
                            {
                                nzp_rod = 568;
                                break;
                            }
                            case "гр.муж":
                            {
                                nzp_rod = 571;
                                break;
                            }
                            case "двоюродн.":
                            {
                                nzp_rod = 15;
                                break;
                            }
                            case "дочь":
                            {
                                nzp_rod = 559;
                                break;
                            }
                            case "дядя":
                            {
                                nzp_rod = 738;
                                break;
                            }
                            case "жена":
                            {
                                nzp_rod = 562;
                                break;
                            }
                            case "зять":
                            {
                                nzp_rod = 565;
                                break;
                            }
                            case "кс":
                            {
                                nzp_rod = 561;
                                break;
                            }
                            case "мать":
                            {
                                nzp_rod = 563;
                                break;
                            }
                            case "мать мужа":
                            {
                                nzp_rod = 619;
                                break;
                            }
                            case "муж":
                            {
                                nzp_rod = 567;
                                break;
                            }
                            case "отец":
                            {
                                nzp_rod = 572;
                                break;
                            }
                            case "отчим":
                            {
                                nzp_rod = 640;
                                break;
                            }
                            case "падчерица":
                            {
                                nzp_rod = 30;
                                break;
                            }
                            case "племянник":
                            {
                                nzp_rod = 666;
                                break;
                            }
                            case "племянница":
                            {
                                nzp_rod = 899;
                                break;
                            }
                            case "сестра":
                            {
                                nzp_rod = 899;
                                break;
                            }
                            case "сноха":
                            {
                                nzp_rod = 575;
                                break;
                            }
                            case "сын":
                            {
                                nzp_rod = 564;
                                break;
                            }
                            case "сын жены":
                            {
                                nzp_rod = 1156;
                                break;
                            }
                            case "тетя":
                            {
                                nzp_rod = 596;
                                break;
                            }
                            case "теща":
                            {
                                nzp_rod = 594;
                                break;
                            }
                        }
                        #endregion
                        int nzp_dok = 0;
                        #region nzp_dok
                        switch (Convert.ToString(book.Worksheet(1).Row(i).Cell(10).Value).Trim())
                        {
                            case "паспорт":
                                {
                                    nzp_dok = 10;
                                    break;
                                }
                            case "Св-во о рожд.":
                                {
                                    nzp_dok = 2;
                                    break;
                                }
                            case "Св-во о рождении":
                                {
                                    nzp_dok = 2;
                                    break;
                                }
                            case "Св-во рожд.":
                                {
                                    nzp_dok = 2;
                                    break;
                                }
                            default:
                                {
                                    nzp_dok = -1;
                                    break;
                                }
                        }
                        #endregion

                        string serij = "";
                        if (Convert.ToString(book.Worksheet(1).Row(i).Cell(11).Value).Trim() != "" && Convert.ToString(book.Worksheet(1).Row(i).Cell(11).Value).Trim().Length >= 4)
                        {
                            if (nzp_dok == 10)
                                serij = Convert.ToString(book.Worksheet(1).Row(i).Cell(11).Value).Trim().Substring(0, 2) + " " + Convert.ToString(book.Worksheet(1).Row(i).Cell(11).Value).Trim().Substring(2, 2);
                            else
                                serij = Convert.ToString(book.Worksheet(1).Row(i).Cell(11).Value).Trim();
                        }
                        int nzp_kart = pg.InsertKart(nzp_gil, nzp_kvar.Split('|')[0],
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(5).Value).Trim().ToUpper(),
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(6).Value).Trim().ToUpper(),
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(7).Value).Trim().ToUpper(),
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(8).Value).Trim(),
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(9).Value).Trim(),
                                                    nzp_dok,
                                                    serij,
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(12).Value).Trim(),
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(13).Value).Trim(),
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(14).Value).Trim(),
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(17).Value).Trim(),
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(18).Value).Trim(),
                                                    nzp_rod,
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(20).Value).Trim());
                        pg.InsertGrgd(nzp_kart);
                    }
                }
                book.Save();
            }
            #endregion

            //Эксель по Инспекторам
            #region 66
            else if (type == 66)
            {
                var book = new XLWorkbook(@"C:\temp\шаблон отчета.xlsx");
                DataTable insp = ora.SelectInspection();
                int row = 2;
                int num = 1;
                for (int i = 0; i < insp.Rows.Count; i++)
                {
                    book.Worksheet(1).Row(i + row).Cell(1).Value = num;
                    num++;
                    book.Worksheet(1).Row(i + row).Cell(2).Value = insp.Rows[i][0].ToString();
                    book.Worksheet(1).Row(i + row).Cell(3).Value = insp.Rows[i][1].ToString();
                    book.Worksheet(1).Row(i + row).Cell(4).Value = insp.Rows[i][2].ToString();
                    book.Worksheet(1).Row(i + row).Cell(5).Value = insp.Rows[i][3].ToString();
                    book.Worksheet(1).Row(i + row).Cell(6).Value = insp.Rows[i][4].ToString();
                    book.Worksheet(1).Row(i + row).Cell(7).Value = insp.Rows[i][5].ToString();
                    book.Worksheet(1).Row(i + row).Cell(8).Value = insp.Rows[i][6].ToString();
                    book.Worksheet(1).Row(i + row).Cell(9).Value = insp.Rows[i][7].ToString();
                }
                book.Save();
            }
            #endregion

            //чтение из txt файла
            #region 67
            else if (type == 67)
            {
                var book = new XLWorkbook(@"C:\temp\report.xlsx");
                string[] lines = System.IO.File.ReadAllLines(@"C:\temp\report2.txt", Encoding.Default);
                int row = 2;
                foreach (string line in lines)
                {
                    book.Worksheet(1).Row(row).Cell(1).Value = line.Split(';')[0];
                    book.Worksheet(1).Row(row).Cell(2).Value = line.Split(';')[1];
                    book.Worksheet(1).Row(row).Cell(3).Value = line.Split(';')[2];
                    book.Worksheet(1).Row(row).Cell(4).Value = line.Split(';')[3];
                    book.Worksheet(1).Row(row).Cell(5).Value = line.Split(';')[4];
                    book.Worksheet(1).Row(row).Cell(6).Value = line.Split(';')[5];
                    row++;
                }
                book.Save();
            }
            #endregion

            //перекидка
            #region 68
            else if (type == 68)
            {
                Console.Write("Введите наименование базы:");
                string database = Console.ReadLine();
                Console.Write("Введите наименование банка:");
                string bank = Console.ReadLine();
                Console.Write("Введите год перерасчета:");
                string year = Console.ReadLine();
                Console.Write("Введите месяц перерасчета:");
                string month = Console.ReadLine();
                var book = new XLWorkbook();
                string comment;
                book = new XLWorkbook(@"C:\temp\Недопоставка по Кр.Коммунаров 17 (изолированные).xlsx");
                comment = "заварен мусоропровод";
                for (int i = 4; i <= 60; i++)
                {
                    List<string> nzp_kvar = pg.SelectNzpKvar(database, Convert.ToString(book.Worksheet(1).Row(i).Cell(3).Value).Trim(),
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(5).Value).Trim(),
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(6).Value).Trim(), 2);
                    if (nzp_kvar == null)
                    {
                        book.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                    else
                    {
                        int nzp_doc_base = pg.InsertDocBase(database, comment);
                        pg.InsertPerekidka(database,
                            Convert.ToInt32(nzp_kvar[0]),
                            Convert.ToDecimal(book.Worksheet(1).Row(i).Cell(9).Value) * (-1),
                            nzp_doc_base,
                            Convert.ToInt32(nzp_kvar[1]),
                            17,
                            101179,
                            year + "-" + month + "-11", Convert.ToInt32(month), Convert.ToInt32(year));
                    }
                }
                book.Save();
                book = new XLWorkbook(@"C:\temp\Недопоставка по Кр.Коммунаров 17 (коммуналка)-1.xlsx");
                comment = "заварен мусоропровод";
                for (int i = 4; i <= 56; i++)
                {
                    List<string> nzp_kvar = pg.SelectNzpKvar(database, Convert.ToString(book.Worksheet(1).Row(i).Cell(3).Value).Trim(),
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(5).Value).Trim(),
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(6).Value).Trim(), 2);
                    if (nzp_kvar == null)
                    {
                        book.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                    else
                    {
                        int nzp_doc_base = pg.InsertDocBase(database, comment);
                        pg.InsertPerekidka(database, Convert.ToInt32(nzp_kvar[0]),
                            Convert.ToDecimal(book.Worksheet(1).Row(i).Cell(9).Value) * (-1), nzp_doc_base,
                            Convert.ToInt32(nzp_kvar[1]),
                            17, 101179,
                            year + "-" + month + "-11", Convert.ToInt32(month), Convert.ToInt32(year));
                    }
                }
                book.Save();
                book = new XLWorkbook(@"C:\temp\Недопоставка по Печерской 151.xlsx");
                comment = "заварен мусоропровод";
                for (int i = 4; i <= 204; i++)
                {
                    List<string> nzp_kvar = pg.SelectNzpKvar(database, Convert.ToString(book.Worksheet(1).Row(i).Cell(3).Value).Trim(),
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(5).Value).Trim(),
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(6).Value).Trim(), 2);
                    if (nzp_kvar == null)
                    {
                        book.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                    else
                    {
                        int nzp_doc_base = pg.InsertDocBase(database, comment);
                        pg.InsertPerekidka(database, Convert.ToInt32(nzp_kvar[0]),
                            Convert.ToDecimal(book.Worksheet(1).Row(i).Cell(9).Value) * (-1), nzp_doc_base,
                            Convert.ToInt32(nzp_kvar[1]),
                            17, 101179,
                            year + "-" + month + "-11", Convert.ToInt32(month), Convert.ToInt32(year));
                    }
                }
                book.Save();
                book = new XLWorkbook(@"C:\temp\Недопоставка по Гастелло 47.3 (коммуналка).xlsx");
                comment = "заварен мусоропровод";
                for (int i = 4; i <= 58; i++)
                {
                    List<string> nzp_kvar = pg.SelectNzpKvar(database, Convert.ToString(book.Worksheet(1).Row(i).Cell(3).Value).Trim(),
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(5).Value).Trim(),
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(6).Value).Trim(), 2);
                    if (nzp_kvar == null)
                    {
                        book.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                    else
                    {
                        int nzp_doc_base = pg.InsertDocBase(database, comment);
                        pg.InsertPerekidka(database, Convert.ToInt32(nzp_kvar[0]),
                            Convert.ToDecimal(book.Worksheet(1).Row(i).Cell(8).Value) * (-1), nzp_doc_base,
                            Convert.ToInt32(nzp_kvar[1]),
                            17, 101179,
                            year + "-" + month + "-11", Convert.ToInt32(month), Convert.ToInt32(year));
                    }
                }
                book.Save();
                book = new XLWorkbook(@"C:\temp\Недопоставка по Гастелло 47.3 (изолированные).xlsx");
                comment = "заварен мусоропровод";
                for (int i = 4; i <= 60; i++)
                {
                    List<string> nzp_kvar = pg.SelectNzpKvar(database, Convert.ToString(book.Worksheet(1).Row(i).Cell(3).Value).Trim(),
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(5).Value).Trim(),
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(6).Value).Trim(), 2);
                    if (nzp_kvar == null)
                    {
                        book.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                    else
                    {
                        int nzp_doc_base = pg.InsertDocBase(database, comment);
                        pg.InsertPerekidka(database, Convert.ToInt32(nzp_kvar[0]),
                            Convert.ToDecimal(book.Worksheet(1).Row(i).Cell(8).Value) * (-1), nzp_doc_base,
                            Convert.ToInt32(nzp_kvar[1]),
                            17, 101179,
                            year + "-" + month + "-11", Convert.ToInt32(month), Convert.ToInt32(year));
                    }
                }
                book.Save();
                book = new XLWorkbook(@"C:\temp\недопоставка Мальцева_10.xlsx");
                comment = "лифт не работает";
                for (int i = 3; i <= 61; i++)
                {
                    List<string> nzp_kvar = pg.SelectNzpKvar2(database,
                        Convert.ToString(book.Worksheet(1).Row(i).Cell(3).Value).Trim(),
                        Convert.ToString(book.Worksheet(1).Row(i).Cell(5).Value).Trim(),
                        Convert.ToString(book.Worksheet(1).Row(i).Cell(6).Value).Trim(), 2);
                    if (nzp_kvar == null)
                    {
                        book.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                    else
                    {
                        int nzp_doc_base = pg.InsertDocBase(database, comment);
                        pg.InsertPerekidka(database, Convert.ToInt32(nzp_kvar[0]),
                            Convert.ToDecimal(book.Worksheet(1).Row(i).Cell(10).Value) * (-1), nzp_doc_base,
                            Convert.ToInt32(nzp_kvar[1]),
                            17, 101179,
                            year + "-" + month + "-11", Convert.ToInt32(month), Convert.ToInt32(year));
                    }
                }
                book.Save();
                //book = new XLWorkbook(@"C:\temp\кр.ком 17 б.xlsx");
                //comment = "Лифт не работал с 01.09-20.09.2015";
                //for (int i = 4; i <= 114; i++)
                //{
                //    List<string> nzp_kvar = pg.SelectNzpKvar2(Convert.ToString(book.Worksheet(1).Row(i).Cell(3).Value).Trim(),
                //                                   Convert.ToString(book.Worksheet(1).Row(i).Cell(5).Value).Trim(),
                //                                   Convert.ToString(book.Worksheet(1).Row(i).Cell(6).Value).Trim(), 2);
                //    if (nzp_kvar == null)
                //    {
                //        book.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                //    }
                //    else
                //    {
                //        int nzp_doc_base = pg.InsertDocBase(comment);
                //        pg.InsertPerekidka(Convert.ToInt32(nzp_kvar[0]), Convert.ToDecimal(book.Worksheet(1).Row(i).Cell(9).Value) * (-1), nzp_doc_base, Convert.ToInt32(nzp_kvar[1]));
                //    }
                //}
                //book = new XLWorkbook(@"C:\temp\Генератор по Гастелло 47 по начислению ОДН электро.xlsx");
                //comment = "перерасчет";
                //for (int i = 4; i <= 135; i++)
                //{
                //    List<string> nzp_kvar = pg.SelectNzpKvar2(Convert.ToString(book.Worksheet(1).Row(i).Cell(1).Value).Trim(),
                //                                            Convert.ToString(book.Worksheet(1).Row(i).Cell(2).Value).Trim().Split('-')[1].Trim().Split(' ')[6].Trim(),
                //                                            Convert.ToString(book.Worksheet(1).Row(i).Cell(2).Value).Trim().Split('-')[2].Trim().Split(' ')[1].Trim(),
                //                                     2);
                //    if (nzp_kvar == null)
                //    {
                //        book.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                //    }
                //    else
                //    {
                //        int nzp_doc_base = pg.InsertDocBase(comment);
                //        pg.InsertPerekidka4(Convert.ToInt32(nzp_kvar[0]), Convert.ToDecimal(book.Worksheet(1).Row(i).Cell(8).Value) * (-1), nzp_doc_base, Convert.ToInt32(nzp_kvar[1]));
                //    }
                //}
                //book.Save();
                //book = new XLWorkbook(@"C:\temp\генератор по Гикам май.xlsx");
                //comment = "перекидка";
                //for (int i = 4; i <= 38; i++)
                //{
                //    List<string> nzp_kvar = pg.SelectNzpKvar2(Convert.ToString(book.Worksheet(1).Row(i).Cell(1).Value).Trim(),
                //                                            Convert.ToString(book.Worksheet(1).Row(i).Cell(2).Value).Trim().Split('-')[1].Trim().Split(' ')[6].Trim(),
                //                                            Convert.ToString(book.Worksheet(1).Row(i).Cell(2).Value).Trim().Split('-')[2].Trim().Split(' ')[1].Trim(),
                //                                     2);
                //    if (nzp_kvar == null)
                //    {
                //        book.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                //    }
                //    else
                //    {
                //        int nzp_doc_base = pg.InsertDocBase(comment);
                //        pg.InsertPerekidka2(Convert.ToInt32(nzp_kvar[0]), Convert.ToDecimal(book.Worksheet(1).Row(i).Cell(7).Value) * (-1), nzp_doc_base, Convert.ToInt32(nzp_kvar[1]));
                //    }
                //}
                //book.Save();
                //book = new XLWorkbook(@"C:\temp\Гастелло 47 (1).xlsx");
                //comment = "Корректировка отопления 2015 г.";
                //for (int i = 4; i <= 135; i++)
                //{
                //    List<string> nzp_kvar = pg.SelectNzpKvar2(Convert.ToString(book.Worksheet(1).Row(i).Cell(3).Value).Trim(),
                //                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(5).Value).Trim(),
                //                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(6).Value).Trim(), 2);
                //    if (nzp_kvar == null)
                //    {
                //        book.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                //    }
                //    else
                //    {
                //        int nzp_doc_base = pg.InsertDocBase(comment);
                //        pg.InsertPerekidka3(Convert.ToInt32(nzp_kvar[0]), Convert.ToDecimal(book.Worksheet(1).Row(i).Cell(11).Value), nzp_doc_base, Convert.ToInt32(nzp_kvar[1]));
                //    }
                //}
                //book.Save();
                //book = new XLWorkbook(@"C:\temp\Гастелло 17.2 комф.xlsx");
                //comment = "Корректировка отопления 2015 г.";
                //for (int i = 4; i <= 139; i++)
                //{
                //    List<string> nzp_kvar = pg.SelectNzpKvar2(Convert.ToString(book.Worksheet(1).Row(i).Cell(3).Value).Trim(),
                //                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(5).Value).Trim(),
                //                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(6).Value).Trim(), 2);
                //    if (nzp_kvar == null)
                //    {
                //        book.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                //    }
                //    else
                //    {
                //        int nzp_doc_base = pg.InsertDocBase(comment);
                //        pg.InsertPerekidka3(Convert.ToInt32(nzp_kvar[0]), Convert.ToDecimal(book.Worksheet(1).Row(i).Cell(11).Value), nzp_doc_base, Convert.ToInt32(nzp_kvar[1]));
                //    }
                //}
                //book.Save();
                //book = new XLWorkbook(@"C:\temp\Генератор Гастелло 47.3.xlsx");
                //comment = "Корректировка отопления 2015 г.";
                //for (int i = 4; i <= 115; i++)
                //{
                //    List<string> nzp_kvar = pg.SelectNzpKvar2(Convert.ToString(book.Worksheet(1).Row(i).Cell(3).Value).Trim(),
                //                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(5).Value).Trim(),
                //                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(6).Value).Trim(), 2);
                //    if (nzp_kvar == null)
                //    {
                //        book.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                //    }
                //    else
                //    {
                //        int nzp_doc_base = pg.InsertDocBase(comment);
                //        pg.InsertPerekidka3(Convert.ToInt32(nzp_kvar[0]), Convert.ToDecimal(book.Worksheet(1).Row(i).Cell(11).Value), nzp_doc_base, Convert.ToInt32(nzp_kvar[1]));
                //    }
                //}
                //book.Save();
                //book = new XLWorkbook(@"C:\temp\SpLsNach_728269186133.xlsx");
                //comment = "Перерасчет";
                //for (int i = 4; i <= 139; i++)
                //{
                //    List<string> nzp_kvar = pg.SelectNzpKvar(Convert.ToString(book.Worksheet(1).Row(i).Cell(1).Value).Trim(),
                //                                            Convert.ToString(book.Worksheet(1).Row(i).Cell(2).Value).Trim().Split('-')[1].Trim().Split(' ')[6].Trim(),
                //                                           Convert.ToString(book.Worksheet(1).Row(i).Cell(2).Value).Trim().Split('-')[2].Trim().Split(' ')[1].Trim(),
                //                                     2);
                //    if (nzp_kvar == null)
                //    {
                //        book.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                //    }
                //    else
                //    {
                //        int nzp_doc_base = pg.InsertDocBase(comment);
                //        pg.InsertPerekidka5(Convert.ToInt32(nzp_kvar[0]), Convert.ToDecimal(book.Worksheet(1).Row(i).Cell(4).Value), nzp_doc_base, Convert.ToInt32(nzp_kvar[1]));
                //    }
                //}
                //book.Save();
                //book = new XLWorkbook(@"C:\temp\Реестр с расхождением в датах посчитанный.xlsx");
                //comment = "Перерасчет";
                //for (int i = 2; i <= 3; i++)
                //{
                //    List<string> nzp_kvar = pg.SelectNzpKvarPkod(Convert.ToString(book.Worksheet(1).Row(i).Cell(1).Value).Trim());
                //    if (nzp_kvar == null)
                //    {
                //        book.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                //    }
                //    else
                //    {
                //        int nzp_doc_base = pg.InsertDocBase(comment);
                //        pg.InsertPerekidka6(Convert.ToInt32(nzp_kvar[0]), Convert.ToDecimal(book.Worksheet(1).Row(i).Cell(2).Value), nzp_doc_base, Convert.ToInt32(nzp_kvar[1]));
                //    }
                //}
                //book.Save();
                //book = new XLWorkbook(@"C:\temp\кр.ком 17 б.xlsx");
                //comment = "Лифт не работал с 18-31.08";
                //for (int i = 4; i <= 114; i++)
                //{
                //    List<string> nzp_kvar = pg.SelectNzpKvar2(Convert.ToString(book.Worksheet(1).Row(i).Cell(3).Value).Trim(),
                //                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(5).Value).Trim(),
                //                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(6).Value).Trim(), 2);
                //    if (nzp_kvar == null)
                //    {
                //        book.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                //    }
                //    else
                //    {
                //        int nzp_doc_base = pg.InsertDocBase(comment);
                //        pg.InsertPerekidka7(Convert.ToInt32(nzp_kvar[0]), Convert.ToDecimal(book.Worksheet(1).Row(i).Cell(10).Value) * (-1), nzp_doc_base, Convert.ToInt32(nzp_kvar[1]));
                //    }
                //}
                //book.Save();
                //book = new XLWorkbook(@"C:\temp\22Парт1А.xlsx");
                //comment = "Не соответствовал температурный режим с 17.09-30.09.2015 г.";
                //for (int i = 2; i <= 83; i++)
                //{
                //    List<string> nzp_kvar = pg.SelectNzpKvarByPkod(Convert.ToString(book.Worksheet(1).Row(i).Cell(1).Value).Trim());
                //    if (nzp_kvar == null)
                //    {
                //        book.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                //    }
                //    else
                //    {
                //        int nzp_doc_base = pg.InsertDocBase(comment);
                //        pg.InsertPerekidka14(Convert.ToInt32(nzp_kvar[0]), Convert.ToDecimal(book.Worksheet(1).Row(i).Cell(2).Value) * (-1), nzp_doc_base, Convert.ToInt32(nzp_kvar[1]));
                //    }
                //}
                //book.Save();
                //book = new XLWorkbook(@"C:\temp\Перерасчет 50.xlsx");
                //comment = "Перерасчет по ИПУ";
                //for (int i = 2; i <= 116; i++)
                //{
                //    int month = 9;
                //    List<string> nzp_kvar = pg.SelectNzpKvarByNumLs("billTlt", Convert.ToString(book.Worksheet(1).Row(i).Cell(1).Value).Trim().Substring(5));
                //    if (nzp_kvar == null)
                //    {
                //        book.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                //    }
                //    else
                //    {
                //        int nzp_doc_base = pg.InsertDocBase("billTlt", Convert.ToString(book.Worksheet(1).Row(i).Cell(2).Value).Trim() != "26" ? comment : "");
                //        pg.InsertPerekidkaByNzpServAndMonthAndSupp("billTlt",
                //                                                        Convert.ToInt32(nzp_kvar[0]), 
                //                                                            Convert.ToDecimal(book.Worksheet(1).Row(i).Cell(3).Value),
                //                                                                nzp_doc_base, 
                //                                                                    Convert.ToInt32(nzp_kvar[1]),
                //                                                                        Convert.ToString(book.Worksheet(1).Row(i).Cell(2).Value).Trim(),
                //                                                                            month,
                //                                                                                Convert.ToString(book.Worksheet(1).Row(i).Cell(4).Value).Trim());
                //    }
                //}
                //book.Save();

                //book = new XLWorkbook(@"C:\temp\Недопоставка за 10.2015 по 22 Парт 1 А по ГВС.xlsx");
                //comment = "Несоответствие температурного режима с 1.10.15-11.10.2015 гг";
                //for (int i = 10; i <= 103; i++)
                //{
                //    int month = 10;
                //    List<string> nzp_kvar = pg.SelectNzpKvarByNumLs(database, Convert.ToString(book.Worksheet(1).Row(i).Cell(4).Value).Trim());
                //    if (nzp_kvar == null)
                //    {
                //        book.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                //    }
                //    else
                //    {
                //        if (book.Worksheet(1).Row(i).Cell(9).Value != "")
                //        {
                //            int nzp_doc_base = pg.InsertDocBase(database, comment);
                //            pg.InsertPerekidkaByNzpServAndMonthAndSupp(database,
                //                                                            Convert.ToInt32(nzp_kvar[0]),
                //                                                                Convert.ToDecimal(book.Worksheet(1).Row(i).Cell(9).Value) * (-1),
                //                                                                    nzp_doc_base,
                //                                                                        Convert.ToInt32(nzp_kvar[1]),
                //                                                                            "9",
                //                                                                                month,
                //                                                                                    "101185");
                //        }
                //    }
                //}
                //book.Save();
                //book = new XLWorkbook(@"C:\temp\Корректировка отопл за 2015 по д.50 правильный.xlsx");
                //comment = "Корректировка по отоплению";
                //for (int i = 2; i <= 263; i++)
                //{
                //    int month = 1;
                //    int year = 2016;
                //    int nzp_dom = 7155107;
                //    List<string> nzp_kvar = pg.SelectNzpKvarByPkod10NzpDomNKvar(database,
                //        Convert.ToString(book.Worksheet(1).Row(i).Cell(3).Value).Trim(), nzp_dom, bank,
                //        Convert.ToString(book.Worksheet(1).Row(i).Cell(6).Value).Trim());

                //    if (nzp_kvar == null)
                //    {
                //        book.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                //    }
                //    else
                //    {
                //        if (book.Worksheet(1).Row(i).Cell(10).Value != "")
                //        {
                //            int nzp_doc_base = pg.InsertDocBase(database, comment);
                //            pg.InsertPerekidkaByNzpServAndMonthAndSupp(database,
                //                                                            Convert.ToInt32(nzp_kvar[0]),
                //                                                                Convert.ToDecimal(book.Worksheet(1).Row(i).Cell(10).Value),
                //                                                                    nzp_doc_base,
                //                                                                        Convert.ToInt32(nzp_kvar[1]),
                //                                                                            "8",
                //                                                                                month,
                //                                                                                    "101191",
                //                                                                                        year,
                //                                                                                            bank);
                //        }
                //    }
                //}
                //book.Save();
            }
            #endregion

            #region 69
            else if (type == 69)
            {
                var book = new XLWorkbook(@"C:\temp\индивидуальныеДома.xlsx");
                for (int i = 4; i <= 7384; i++)//7384
                {
                    decimal total_area = 0;
                    int people_count = 0;
                    if (Convert.ToString(book.Worksheet(1).Row(i).Cell(4).Value).Trim() != "")
                    {
                        total_area = Convert.ToDecimal(book.Worksheet(1).Row(i).Cell(4).Value);
                    }
                    if (Convert.ToString(book.Worksheet(1).Row(i).Cell(5).Value).Trim() != "")
                    {
                        people_count = Convert.ToInt32(book.Worksheet(1).Row(i).Cell(5).Value);
                    }
                    string str = ora.InsertRealtyObject(Convert.ToString(book.Worksheet(1).Row(i).Cell(2).Value),
                                                        Convert.ToString(book.Worksheet(1).Row(i).Cell(3).Value),
                                                        total_area, people_count);
                    if (str != "ЗАГРУЖЕНО")
                    {
                        book.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                        book.Worksheet(1).Row(i).Cell(7).Value = str;
                    }
                    if(i%500 == 0)
                        Console.WriteLine(i.ToString());
                }
            }
            #endregion

             //тест ЕЛСЕ
            #region 70
            else if (type == 70)
            {
                int i = 0;
                string str = (i > 0) ? "Готово!!!" : "Готово!!!";
            }
            #endregion

            //тест ЕЛСЕ
            #region 701
            else if (type == 701)
            {
                Console.WriteLine(Convert.ToDecimal("0,0379"));
            }
            #endregion

            //Загрузка платежей в стройку
            #region 71
            else if (type == 71)
            {
                int nameCount = 1;
                int docCount = 1;
                var wb2 = new XLWorkbook(@"C:\temp\GosContract.xlsx");
                for (int i = 3; i <= 3203; i++)
                {
                    if (i % 100 == 0)
                        Console.WriteLine(i.ToString());
                    DateTime docDate;
                    bool result = DateTime.TryParse(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim(), out docDate);
                    if (!result)
                    {
                        docDate = new DateTime(2001,1,1);
                    }

                    DateTime termsOfObligation;
                    bool result2 = DateTime.TryParse(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim(), out termsOfObligation);
                    if (!result2)
                    {
                        termsOfObligation = new DateTime(1111, 1, 1);
                    }
                    string docNumber = "";
                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value) != null && Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value).Trim() != "")
                        docNumber = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value).Trim();
                    if (docNumber.Trim() == "")
                    {
                        docNumber = "не указано " + docCount.ToString();
                        docCount++;
                    }
                    int recipientId;
                    bool result3 = Int32.TryParse(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(11).Value).Trim(), out recipientId);
                    if (!result3)
                    {
                        string part1 = "";
                        string part2 = "";
                        string name;
                        if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(10).Value) != null && Convert.ToString(wb2.Worksheet(1).Row(i).Cell(10).Value).Trim() != "")
                            part1 = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(10).Value).Trim();
                        if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(11).Value) != null && Convert.ToString(wb2.Worksheet(1).Row(i).Cell(11).Value).Trim() != "")
                            part2 = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(11).Value).Trim();
                        name = part1 + " " + part2;
                        if (name.Trim() == "")
                        {
                            name = "не указано " + nameCount.ToString();
                            nameCount++;
                        }
                        int str = pg.InsertRecipientID(name);
                        if (str == 0)
                        {
                            recipientId = pg.InsertRecipient(name.Trim());
                            if (recipientId == 0)
                            {
                                Console.WriteLine("ошибка при добавлении поставщика: строка = " + i.ToString());
                                break;
                            }
                        }
                        else
                        {
                            recipientId = str;
                        }
                    }

                    decimal amount;
                    bool result4 = Decimal.TryParse(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(14).Value).Trim(), out amount);
                    if (!result4)
                    {
                        amount = 0;
                    }

                    int gosContractId = pg.InsertGosContract(docNumber, docDate, recipientId, amount.ToString().Replace(',', '.'), termsOfObligation);
                    if (gosContractId == 0)
                    {
                        wb2.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Blue;
                    }
                    else if (gosContractId == 1)
                    {
                        Console.WriteLine("ошибка при добавлении гос контракта: строка = " + i.ToString());
                        break;
                    }
                    else
                    {
                        int object_aip_id;
                        bool result5 = Int32.TryParse(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(12).Value).Trim(), out object_aip_id);
                        if (!result5)
                        {
                            object_aip_id = 0;
                        }
                        int gosContractRec = pg.InsertGosContractRec(gosContractId, object_aip_id, amount.ToString().Replace(',', '.'));
                        if (gosContractRec == 0)
                        {
                            Console.WriteLine("ошибка при добавлении гос контракт рекорда: строка = " + i.ToString());
                            break;
                        }
                    }

                }
                wb2.Save();
            }
            #endregion

            #region 72
            else if (type == 72)
            {
                var wb2 = new XLWorkbook(@"C:\temp\GosContract.xlsx");
                for (int i = 2; i <= 3203; i++)
                {
                    if (i % 200 == 0)
                        Console.WriteLine(i.ToString());
                    string str = pg.InsertRecipientID(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(10).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(11).Value).Trim());
                    if (str.Split('|')[0] != "ЗАГРУЖЕНО")
                    {
                        wb2.Worksheet(1).Row(i).Cell(11).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(1).Row(i).Cell(20).Value = str.Split('|')[1];
                    }
                    else
                    {
                        wb2.Worksheet(1).Row(i).Cell(11).Value = str.Split('|')[1];
                        wb2.Worksheet(1).Row(i).Cell(20).Value = str.Split('|')[2];
                    }
                }
                wb2.Save();
            }
            #endregion
            #region 73
            else if (type == 73)
            {
                var wb2 = new XLWorkbook(@"C:\temp\GosContract.xlsx");
                for (int i = 2; i <= 3203; i++)
                {
                    if (i % 200 == 0)
                        Console.WriteLine(i.ToString());
                    int recipientId;
                    bool result3 = Int32.TryParse(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(11).Value).Trim(), out recipientId);
                    if (result3)
                    {
                        wb2.Worksheet(1).Row(i).Cell(11).Style.Fill.BackgroundColor = XLColor.Orange;
                    }
                }
                wb2.Save();
            }
            #endregion
            #region 74
            else if (type == 74)
            {
                var wb2 = new XLWorkbook(@"C:\temp\GosContract.xlsx");
                for (int i = 2; i <= 3203; i++)
                {
                    if (i % 200 == 0)
                        Console.WriteLine(i.ToString());
                    int obkectId;
                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(12).Value) != null && Convert.ToString(wb2.Worksheet(1).Row(i).Cell(12).Value).Trim() != "")
                    {
                        int objId;
                        bool result3 = Int32.TryParse(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(12).Value).Trim(), out objId);
                        if (!result3)
                        {
                            obkectId = pg.GetObjectId(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(12).Value).Trim());
                            if (obkectId == 0)
                                wb2.Worksheet(1).Row(i).Cell(12).Style.Fill.BackgroundColor = XLColor.Yellow;
                            else if (obkectId == 1)
                                wb2.Worksheet(1).Row(i).Cell(12).Style.Fill.BackgroundColor = XLColor.Blue;
                            else if (obkectId == 2)
                                wb2.Worksheet(1).Row(i).Cell(12).Style.Fill.BackgroundColor = XLColor.Green;
                            else
                                wb2.Worksheet(1).Row(i).Cell(12).Value = obkectId;
                        }
                    }

                }
                wb2.Save();
            }
            #endregion

            #region 75
            else if (type == 75)
            {
                DataTable dtHouse = new DataTable();
                DataTable dtPeople = new DataTable();
                var wb2 = new XLWorkbook(@"C:\temp\Сведения по Абонентам.xlsx");
                StreamWriter sw = new StreamWriter(@"C:\temp\peopleError.txt", false);
                for (int i = 6; i <= 71373; i++)
                {
                    try
                    {
                        int flat_number;
                        string flat = "";
                        bool result2 = Int32.TryParse(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value).Trim(), out flat_number);
                        if (!result2)
                        {
                            flat = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value);
                        }
                        else
                        {
                            flat = flat_number.ToString();
                        }

                        int res_count;
                        bool result = Int32.TryParse(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(9).Value).Trim(), out res_count);
                        if (!result)
                        {
                            res_count = 0;
                        }
                        string str = ora.InsertPeople3(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value),
                            flat, 
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value),
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(8).Value),
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(7).Value),
                            res_count.ToString(),
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value));
                        if (str != "ЗАГРУЖЕНО")
                        {
                            Console.WriteLine(str);
                            sw.WriteLine(i.ToString() + " = " + str);
                        }
                    }
                    catch
                    {

                    }

                }
                sw.Close();
                wb2.Save();
            }
            #endregion

            #region 76
            else if (type == 76)
            {
                var wb2 = new XLWorkbook(@"C:\temp\запорожская.39 цифры-1.xlsx");
                for (int i = 6; i <= 143; i++)//143
                {
                    string nzp_kvar = pg.SelectNzpKvar2(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value));
                    if (nzp_kvar.Split('|')[1] == "Найдено")
                    {
                        pg.UpdateSaldo(Math.Round(Convert.ToDecimal(wb2.Worksheet(1).Row(i).Cell(12).Value), 4),
                            Math.Round(Convert.ToDecimal(wb2.Worksheet(1).Row(i).Cell(13).Value), 4),
                            Math.Round(Convert.ToDecimal(wb2.Worksheet(1).Row(i).Cell(16).Value) , 4), 
                            Convert.ToInt32(nzp_kvar.Split('|')[0]),
                            Math.Round(Convert.ToDecimal(wb2.Worksheet(1).Row(i).Cell(17).Value) , 4), 
                            Math.Round(Convert.ToDecimal(wb2.Worksheet(1).Row(i).Cell(18).Value) , 4));
                    }
                    else
                    {
                        Console.WriteLine(i.ToString() + " = " + nzp_kvar);
                        break;
                    }
                }
            }
            #endregion



            #region 77
            else if (type == 77)
            {
                DataTable dtHouse = new DataTable();
                DataTable dtPeople = new DataTable();
                var wb2 = new XLWorkbook(@"C:\temp\Копия Копия Сведения о квартирах (реестр)-2-1.xlsx");
                string address = "";
                for (int j = 1; j <= 10; j++)
                {
                    Console.WriteLine(j.ToString());
                    for (int i = 3; i <= 1000; i++)
                    {
                        try
                        {
                            if (Convert.ToString(wb2.Worksheet(j).Row(i).Cell(1).Value) == "")
                            {
                                
                            }
                            else
                            {

                                string str = ora.InsertPeople2(Convert.ToString(wb2.Worksheet(j).Row(1).Cell(1).Value).Trim(),
                                    (j != 10) ? Convert.ToInt32(Convert.ToString(wb2.Worksheet(j).Row(i).Cell(1).Value).Trim().Substring(1)) : Convert.ToInt32(Convert.ToString(wb2.Worksheet(j).Row(i).Cell(1).Value)),
                                    Convert.ToString(wb2.Worksheet(j).Row(i).Cell(6).Value).Trim(),
                                    Convert.ToString(wb2.Worksheet(j).Row(i).Cell(2).Value),
                                    Convert.ToString(wb2.Worksheet(j).Row(i).Cell(3).Value),
                                    Convert.ToString(wb2.Worksheet(j).Row(i).Cell(5).Value),
                                    Convert.ToString(wb2.Worksheet(j).Row(i).Cell(4).Value));
                                if (str != "ЗАГРУЖЕНО")
                                    Console.WriteLine(str);
                            }
                        }
                        catch(Exception e)
                        {
                            Console.WriteLine(i.ToString() + " - " + e.ToString());
                        }

                    }
                }
                wb2.Save();
            }
            #endregion

            #region 78 Мельникову
            else if (type == 78)
            {
                DataTable obj = ora.SelectDisp();
                var wb = new XLWorkbook();
                wb.AddWorksheet("report");
                string manorg = obj.Rows[0][0].ToString();
                string type_ = obj.Rows[0][1].ToString();
                int rowMove = 1;
                for (int i = 0; i < obj.Rows.Count; i++)
                {
                    if (manorg != obj.Rows[i][0].ToString() || type_ != obj.Rows[i][1].ToString())
                    {
                        rowMove++;
                        wb.Worksheet(1).Row(rowMove).Cell(1).Value = obj.Rows[i][1].ToString();
                        wb.Worksheet(1).Row(rowMove).Cell(2).Value = obj.Rows[i][2].ToString();
                        wb.Worksheet(1).Row(rowMove).Cell(3).Value = obj.Rows[i][4].ToString();
                        wb.Worksheet(1).Row(rowMove).Cell(4).Value = obj.Rows[i][5].ToString();
                        manorg = obj.Rows[i][0].ToString();
                        type_ = obj.Rows[i][1].ToString();
                    }
                    else
                    {
                        rowMove++;
                        wb.Worksheet(1).Row(rowMove).Cell(1).Value = "Проверка исполнений предписаний";
                        wb.Worksheet(1).Row(rowMove).Cell(2).Value = obj.Rows[i][2].ToString();
                        wb.Worksheet(1).Row(rowMove).Cell(3).Value = obj.Rows[i][4].ToString();
                        wb.Worksheet(1).Row(rowMove).Cell(4).Value = obj.Rows[i][5].ToString();
                    }
                }
                wb.SaveAs(@"C:\temp\report" + DateTime.Now.ToShortDateString() + ".xlsx");
            }
            #endregion


            //ftp connect
            #region 79
            else if (type == 79)
            {
                FtpClient client = new FtpClient();
                //Задаём параметры клиента.
                client.PassiveMode = true; //Включаем пассивный режим.
                int TimeoutFTP = 30000; //Таймаут.
                string FTP_SERVER = "ftp.zakupki.gov.ru";
                //Подключаемся к FTP серверу.
                client.Connect(TimeoutFTP, FTP_SERVER, 21);
                client.Login(TimeoutFTP, "free", "free");
                client.ChangeDirectory(TimeoutFTP, "fcs_regions/Samarskaja_obl/contracts");
                string pathContract = @"C:\temp\depstr\contracts\incoming";
                string pathContractExtract = @"C:\temp\depstr\contracts\extract";
                string pathContractFileLoad = @"C:\temp\depstr\contracts\fileLoad";
                int un = 0;
                Directory.SetCurrentDirectory(pathContract);
                /*if (Directory.Exists(pathContract + @"\" + DateTime.Today.ToShortDateString()))
                {
                    Directory.SetCurrentDirectory(pathContract);
                    Directory.Delete(Directory.GetCurrentDirectory() + @"\" + DateTime.Today.ToShortDateString(), true);
                    Directory.CreateDirectory(DateTime.Today.ToShortDateString());
                    Directory.SetCurrentDirectory(pathContract + @"\" + DateTime.Today.ToShortDateString());
                }
                else
                {
                    Directory.CreateDirectory(DateTime.Today.ToShortDateString());
                    Directory.SetCurrentDirectory(pathContract + @"\" + DateTime.Today.ToShortDateString());
                }*/
                foreach (var t in client.GetDirectoryList(TimeoutFTP))
                {
                    if (!File.Exists(Directory.GetCurrentDirectory() + @"\" + t.Name))
                    {
                        string file = Directory.GetCurrentDirectory() + @"\" + t.Name;
                        client.GetFile(TimeoutFTP, file, t.Name);
                        //C:\Temp\7-Zip
                        //ZipFile.ExtractToDirectory(file, pathContractExtract);
                        // Формируем параметры вызова 7z
                        ProcessStartInfo startInfo = new ProcessStartInfo();
                        startInfo.FileName = @"C:\Temp\7-Zip\7z.exe";
                        // Распаковать (для полных путей - x)
                        startInfo.Arguments = " e";
                        // На все отвечать yes
                        startInfo.Arguments += " -y";
                        // Файл, который нужно распаковать
                        startInfo.Arguments += " " + "\"" + file + "\"";
                        // Папка распаковки
                        startInfo.Arguments += " -o" + "\"" + pathContractExtract + "\"";
                        startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                        int sevenZipExitCode = 0;
                        using (Process sevenZip = Process.Start(startInfo))
                        {
                            sevenZip.WaitForExit();
                            sevenZipExitCode = sevenZip.ExitCode;
                        }
                        // Если с первого раза не получилось,
                        //пробуем еще раз через 1 секунду
                        if (sevenZipExitCode != 0 && sevenZipExitCode != 1)
                        {
                            using (Process sevenZip = Process.Start(startInfo))
                            {
                                Thread.Sleep(1000);
                                sevenZip.WaitForExit();
                                switch (sevenZip.ExitCode)
                                {
                                    case 0: return; // Без ошибок и предупреждений
                                    case 1: return; // Есть некритичные предупреждения
                                    case 2: throw new Exception("Фатальная ошибка");
                                    case 7: throw new Exception("Ошибка в командной строке");
                                    case 8:
                                    throw new Exception("Недостаточно памяти для выполнения операции");
                                    case 225:
                                    throw new Exception("Пользователь отменил выполнение операции");
                                    default: throw new Exception("Архиватор 7z вернул недокументированный код ошибки: " + sevenZip.ExitCode.ToString());
                                }
                            }
                        }
                        DirectoryInfo dir = new DirectoryInfo(pathContractExtract);
                        foreach (var item in dir.GetFiles())
                        {
                            bool findInn = false;
                            bool findKpp = false;
                            string str = pathContractExtract + @"\" + item.Name;
                            using (XmlReader reader = XmlReader.Create(str))
                            {
                                while (reader.Read())
                                {
                                    string tmp = reader.Name;
                                    if (tmp == "oos:inn" || tmp == "INN")
                                    {
                                        if (tmp == "oos:inn")
                                        {
                                            reader.ReadStartElement("oos:inn");
                                            if (reader.ReadString() == "6315700286")
                                                findInn = true;
                                            //reader.ReadEndElement();
                                        }
                                        else
                                        {
                                            reader.ReadStartElement("INN");
                                            if (reader.ReadString() == "6315700286")
                                                findInn = true;
                                            //reader.ReadEndElement();
                                        }
                                    }
                                    if (tmp == "oos:kpp" || tmp == "KPP")
                                    {
                                        if (tmp == "oos:kpp")
                                        {
                                                reader.ReadStartElement("oos:kpp");
                                                if (reader.ReadString() == "631501001")
                                                    findKpp = true;
                                                //reader.ReadEndElement();
                                        }
                                        else
                                        {
                                            reader.ReadStartElement("KPP");
                                            if (reader.ReadString() == "631501001")
                                                findKpp = true;
                                            //reader.ReadEndElement();
                                        }
                                    }
                                }
                            }
                            if (findKpp && findInn)
                            {
                                if (File.Exists(pathContractFileLoad + @"\" + item.Name))
                                {
                                    File.Copy(str, pathContractFileLoad + @"\"+ un + item.Name);
                                    un++;
                                }
                                else
                                    File.Copy(str, pathContractFileLoad + @"\" + item.Name);
                            }
                        }
                        foreach (var item in dir.GetFiles())
                        {
                            File.Delete( pathContractExtract + @"\" + item.Name);
                        }
                    }
                }
                

                client.Disconnect(TimeoutFTP);
            }
            #endregion

            //ftp connect
            #region 80
            else if (type == 80)
            {
                FtpClient client = new FtpClient();
                //Задаём параметры клиента.
                client.PassiveMode = true; //Включаем пассивный режим.
                int TimeoutFTP = 30000; //Таймаут.
                string FTP_SERVER = "ftp.zakupki.gov.ru";
                //Подключаемся к FTP серверу.
                client.Connect(TimeoutFTP, FTP_SERVER, 21);
                client.Login(TimeoutFTP, "free", "free");
                //client.ChangeDirectory(TimeoutFTP, "fcs_regions/Samarskaja_obl/contracts");
                client.ChangeDirectory(TimeoutFTP, "fcs_regions/Samarskaja_obl/notifications");
                string lotNumber = "0342300000115000147";

                string pathNotifikation = @"C:\temp\depstr\notifications\incoming";


                string pathContract = @"C:\temp\depstr\contracts\incoming";
                string pathContractExtract = @"C:\temp\depstr\contracts\extract";
                string pathContractFileLoad = @"C:\temp\depstr\contracts\fileLoad";
                int un = 0;
                Directory.SetCurrentDirectory(pathNotifikation);
                foreach (var t in client.GetDirectoryList(TimeoutFTP))
                {
                    if (!File.Exists(Directory.GetCurrentDirectory() + @"\" + t.Name) && t.Name.Substring(t.Name.Length - 3) == "zip")
                    {
                        string file = Directory.GetCurrentDirectory() + @"\" + t.Name;
                        client.GetFile(TimeoutFTP, file, t.Name);
                        //C:\Temp\7-Zip
                        //ZipFile.ExtractToDirectory(file, pathContractExtract);
                        // Формируем параметры вызова 7z
                        ProcessStartInfo startInfo = new ProcessStartInfo();
                        startInfo.FileName = @"C:\Temp\7-Zip\7z.exe";
                        // Распаковать (для полных путей - x)
                        startInfo.Arguments = " e";
                        // На все отвечать yes
                        startInfo.Arguments += " -y";
                        // Файл, который нужно распаковать
                        startInfo.Arguments += " " + "\"" + file + "\"";
                        // Папка распаковки
                        startInfo.Arguments += " -o" + "\"" + pathContractExtract + "\"";
                        startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                        int sevenZipExitCode = 0;
                        using (Process sevenZip = Process.Start(startInfo))
                        {
                            sevenZip.WaitForExit();
                            sevenZipExitCode = sevenZip.ExitCode;
                        }
                        // Если с первого раза не получилось,
                        //пробуем еще раз через 1 секунду
                        if (sevenZipExitCode != 0 && sevenZipExitCode != 1)
                        {
                            using (Process sevenZip = Process.Start(startInfo))
                            {
                                Thread.Sleep(1000);
                                sevenZip.WaitForExit();
                                switch (sevenZip.ExitCode)
                                {
                                    case 0: return; // Без ошибок и предупреждений
                                    case 1: return; // Есть некритичные предупреждения
                                    case 2: throw new Exception("Фатальная ошибка");
                                    case 7: throw new Exception("Ошибка в командной строке");
                                    case 8:
                                        throw new Exception("Недостаточно памяти для выполнения операции");
                                    case 225:
                                        throw new Exception("Пользователь отменил выполнение операции");
                                    default: throw new Exception("Архиватор 7z вернул недокументированный код ошибки: " + sevenZip.ExitCode.ToString());
                                }
                            }
                        }
                        DirectoryInfo dir = new DirectoryInfo(pathContractExtract);
                        foreach (var item in dir.GetFiles())
                        {
                            bool findInn = false;
                            bool findKpp = false;
                            string str = pathContractExtract + @"\" + item.Name;
                            using (XmlReader reader = XmlReader.Create(str))
                            {
                                while (reader.Read())
                                {
                                    string tmp = reader.Name;
                                    if (tmp == "oos:inn" || tmp == "INN")
                                    {
                                        if (tmp == "oos:inn")
                                        {
                                            reader.ReadStartElement("oos:inn");
                                            if (reader.ReadString() == "6315700286")
                                                findInn = true;
                                            //reader.ReadEndElement();
                                        }
                                        else
                                        {
                                            reader.ReadStartElement("INN");
                                            if (reader.ReadString() == "6315700286")
                                                findInn = true;
                                            //reader.ReadEndElement();
                                        }
                                    }
                                    if (tmp == "oos:kpp" || tmp == "KPP")
                                    {
                                        if (tmp == "oos:kpp")
                                        {
                                            reader.ReadStartElement("oos:kpp");
                                            if (reader.ReadString() == "631501001")
                                                findKpp = true;
                                            //reader.ReadEndElement();
                                        }
                                        else
                                        {
                                            reader.ReadStartElement("KPP");
                                            if (reader.ReadString() == "631501001")
                                                findKpp = true;
                                            //reader.ReadEndElement();
                                        }
                                    }
                                }
                            }
                            if (findKpp && findInn)
                            {
                                if (File.Exists(pathContractFileLoad + @"\" + item.Name))
                                {
                                    File.Copy(str, pathContractFileLoad + @"\" + un + item.Name);
                                    un++;
                                }
                                else
                                    File.Copy(str, pathContractFileLoad + @"\" + item.Name);
                            }
                        }
                        foreach (var item in dir.GetFiles())
                        {
                            File.Delete(pathContractExtract + @"\" + item.Name);
                        }
                    }
                }


                client.Disconnect(TimeoutFTP);
            }
            #endregion

            #region 83
            else if (type == 83)
            {
                List<string> codes = new List<string> { "ZKP44", "ZKР44", "ZK44", "EAР44", "EAP44", "EA44", "ЕPP44", "ЕРР44", "ЕРP44", "ЕPР44", "ОК44", "ZP44", "ZР44", "ОКP44", "OKP44" };
                StreamWriter sw = new StreamWriter(@"C:\temp\depstr\error.log", false);
                bool firstProtocol = true;
                string lotNumber = "";
                FtpClient client = new FtpClient();
                //Задаём параметры клиента.
                client.PassiveMode = true; //Включаем пассивный режим.
                int TimeoutFTP = 30000; //Таймаут.
                string FTP_SERVER = "ftp.zakupki.gov.ru";
                //Подключаемся к FTP серверу.
                client.Connect(TimeoutFTP, FTP_SERVER, 21);
                client.Login(TimeoutFTP, "free", "free");
                DateTime docPublishDateTemp = new DateTime();
                //client.ChangeDirectory(TimeoutFTP, "fcs_regions/Samarskaja_obl/contracts");
                client.ChangeDirectory(TimeoutFTP, "fcs_regions/Samarskaja_obl/notifications");
                string pathNotifikation = @"C:\temp\depstr\notifications\incoming";
                string pathNotifikationExtract = @"C:\temp\depstr\notifications\extract";
                string pathNotifikationFileLoad = @"C:\temp\depstr\notifications\fileLoad";
                string pathProtocols = @"C:\temp\depstr\protocols\incoming";
                string pathProtocolsExtract = @"C:\temp\depstr\protocols\extract";
                string pathProtocolsFileLoad = @"C:\temp\depstr\protocols\fileLoad";
                DirectoryInfo dirIncoming = new DirectoryInfo(pathNotifikation);
                DirectoryInfo dirIncoming2 = new DirectoryInfo(pathProtocols);
                int un = 0;
                Directory.SetCurrentDirectory(pathNotifikation);
                foreach (var t in client.GetDirectoryList(TimeoutFTP))
                {
                    if (t.Name.Substring(t.Name.Length - 3) == "zip" && t.Name.Contains("2015")
                        && !System.IO.File.Exists(Directory.GetCurrentDirectory() + @"\" + t.Name))
                    {
                        string file = Directory.GetCurrentDirectory() + @"\" + t.Name;
                        try
                        {
                            client.GetFile(TimeoutFTP, file, t.Name);
                        }
                        catch
                        {
                            System.IO.File.Delete(file);
                        }
                    }
                }
                client.Disconnect(TimeoutFTP);
                client.Connect(TimeoutFTP, FTP_SERVER, 21);
                client.Login(TimeoutFTP, "free", "free");
                client.ChangeDirectory(TimeoutFTP, "fcs_regions/Samarskaja_obl/protocols");
                Directory.SetCurrentDirectory(pathProtocols);
                foreach (var t in client.GetDirectoryList(TimeoutFTP))
                {
                    if (t.Name.Substring(t.Name.Length - 3) == "zip" && t.Name.Contains("2015")
                            && !System.IO.File.Exists(Directory.GetCurrentDirectory() + @"\" + t.Name))
                    {
                        string file = Directory.GetCurrentDirectory() + @"\" + t.Name;
                        try
                        {
                            client.GetFile(TimeoutFTP, file, t.Name);
                        }
                        catch
                        {
                            System.IO.File.Delete(file);
                        }
                    }
                }
                client.Disconnect(TimeoutFTP);
                DirectoryInfo dirExtract = new DirectoryInfo(pathNotifikationExtract);
                foreach (var item in dirExtract.GetFiles())
                {
                    System.IO.File.Delete(pathNotifikationExtract + @"\" + item.Name);
                }
                DirectoryInfo dirExtract2 = new DirectoryInfo(pathProtocolsExtract);
                foreach (var item in dirExtract2.GetFiles())
                {
                    System.IO.File.Delete(pathProtocolsExtract + @"\" + item.Name);
                }
                try
                {
                    for (int i = 0; i < 1;i++ )
                    {
                        lotNumber = "0342300000115000118";
                        sw.WriteLine("lotNumber = " + lotNumber);
                        firstProtocol = true;
                        DateTime docPublishDate = new DateTime(9999, 12, 12);
                        foreach (var items in dirIncoming.GetFiles())
                        {
                            sw.WriteLine("file= " + pathNotifikation + @"\" + items.Name);
                            string file = pathNotifikation + @"\" + items.Name;
                            //C:\Temp\7-Zip
                            //ZipFile.ExtractToDirectory(file, pathContractExtract);
                            // Формируем параметры вызова 7z
                            ProcessStartInfo startInfo = new ProcessStartInfo();
                            startInfo.FileName = @"C:\Temp\7-Zip\7z.exe";
                            // Распаковать (для полных путей - x)
                            startInfo.Arguments = " e";
                            // На все отвечать yes
                            startInfo.Arguments += " -y";
                            // Файл, который нужно распаковать
                            startInfo.Arguments += " " + "\"" + file + "\"";
                            // Папка распаковки
                            startInfo.Arguments += " -o" + "\"" + pathNotifikationExtract + "\"";
                            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                            int sevenZipExitCode = 0;
                            using (Process sevenZip = Process.Start(startInfo))
                            {
                                sevenZip.WaitForExit();
                                sevenZipExitCode = sevenZip.ExitCode;
                            }
                            // Если с первого раза не получилось,
                            //пробуем еще раз через 1 секунду
                            if (sevenZipExitCode != 0 && sevenZipExitCode != 1)
                            {
                                using (Process sevenZip = Process.Start(startInfo))
                                {
                                    Thread.Sleep(1000);
                                    sevenZip.WaitForExit();
                                }
                            }

                            foreach (var item in dirExtract.GetFiles())
                            {
                                if (item.Name.Contains("Notification"))
                                {
                                    string purchaseObjectInfo = "";
                                    decimal maxPrice = 0;
                                    string code = "";
                                    bool isWrite = false;
                                    bool boolOpening = false;
                                    bool boolScoring = false;
                                    bool writeOpening = false;
                                    bool writeScoring = false;
                                    DateTime grantStartDate = new DateTime();
                                    DateTime grantEndDate = new DateTime();
                                    DateTime opening = new DateTime();
                                    DateTime scoring = new DateTime();

                                    string str = pathNotifikationExtract + @"\" + item.Name;
                                    using (XmlReader reader = XmlReader.Create(str))
                                    {
                                        while (reader.Read())
                                        {
                                            string tmp = reader.Name;
                                            if (tmp == "purchaseObjectInfo")
                                            {
                                                purchaseObjectInfo = reader.ReadString();
                                            }
                                            if (tmp == "opening" && !writeOpening)
                                            {
                                                boolOpening = true;
                                                writeOpening = true;
                                            }
                                            if (tmp == "date" && boolOpening)
                                            {
                                                opening = Convert.ToDateTime(reader.ReadString());
                                                boolOpening = false;
                                            }
                                            if (tmp == "scoring" && !writeScoring)
                                            {
                                                boolScoring = true;
                                                writeScoring = true;
                                            }
                                            if (tmp == "date" && boolScoring)
                                            {
                                                boolScoring = false;
                                                scoring = Convert.ToDateTime(reader.ReadString());
                                            }
                                            if (tmp == "grantStartDate")
                                            {
                                                grantStartDate = Convert.ToDateTime(reader.ReadString());
                                            }
                                            if (tmp == "grantEndDate")
                                            {
                                                grantEndDate = Convert.ToDateTime(reader.ReadString());
                                            }
                                            if (tmp == "maxPrice")
                                            {
                                                string price = reader.ReadString();
                                                try
                                                {
                                                    maxPrice = Convert.ToDecimal(price);
                                                }
                                                catch
                                                {
                                                    maxPrice = Convert.ToDecimal(price.Replace('.', ','));
                                                }
                                            }
                                            if (tmp == "code")
                                            {
                                                string tempCode = reader.ReadString();
                                                if (codes.Contains(tempCode))
                                                    code = tempCode;
                                            }
                                            if (tmp == "purchaseNumber")
                                            {
                                                if (reader.ReadString() == lotNumber)
                                                    isWrite = true;
                                            }
                                            if (tmp == "docPublishDate")
                                            {
                                                docPublishDateTemp = Convert.ToDateTime(reader.ReadString());
                                            }
                                        }
                                    }
                                    if (isWrite)
                                    {
                                        if (System.IO.File.Exists(pathNotifikationFileLoad + @"\" + item.Name))
                                        {
                                            System.IO.File.Copy(str, pathNotifikationFileLoad + @"\" + un + item.Name);
                                            un++;
                                        }
                                        else
                                            System.IO.File.Copy(str, pathNotifikationFileLoad + @"\" + item.Name);
                                    }
                                }
                            }
                            foreach (var item in dirExtract.GetFiles())
                            {
                                System.IO.File.Delete(pathNotifikationExtract + @"\" + item.Name);
                            }
                        }
                        foreach (var items in dirIncoming2.GetFiles())
                        {
                            string file = pathProtocols + @"\" + items.Name;
                            sw.WriteLine("file= " + pathProtocols + @"\" + items.Name);
                            //C:\Temp\7-Zip
                            //ZipFile.ExtractToDirectory(file, pathContractExtract);
                            // Формируем параметры вызова 7z
                            ProcessStartInfo startInfo = new ProcessStartInfo();
                            startInfo.FileName = @"C:\Temp\7-Zip\7z.exe";
                            // Распаковать (для полных путей - x)
                            startInfo.Arguments = " e";
                            // На все отвечать yes
                            startInfo.Arguments += " -y";
                            // Файл, который нужно распаковать
                            startInfo.Arguments += " " + "\"" + file + "\"";
                            // Папка распаковки
                            startInfo.Arguments += " -o" + "\"" + pathProtocolsExtract + "\"";
                            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                            int sevenZipExitCode = 0;
                            using (Process sevenZip = Process.Start(startInfo))
                            {
                                sevenZip.WaitForExit();
                                sevenZipExitCode = sevenZip.ExitCode;
                            }
                            // Если с первого раза не получилось,
                            //пробуем еще раз через 1 секунду
                            if (sevenZipExitCode != 0 && sevenZipExitCode != 1)
                            {
                                using (Process sevenZip = Process.Start(startInfo))
                                {
                                    Thread.Sleep(1000);
                                    sevenZip.WaitForExit();
                                }
                            }
                            foreach (var item in dirExtract2.GetFiles())
                            {
                                if (item.Name.Contains("Protocol"))
                                {
                                    Dictionary<string, decimal> organizationName = new Dictionary<string, decimal>();
                                    List<string> winners = new List<string>();
                                    string orgName = "";
                                    string inn = "";
                                    string purchaseObjectInfo = "";
                                    bool boolOffer = false;
                                    bool isWrite = false;
                                    DateTime protocolDate = new DateTime();
                                    string str = pathProtocolsExtract + @"\" + item.Name;
                                    if (firstProtocol)
                                    {
                                        using (XmlReader reader = XmlReader.Create(str))
                                        {
                                            while (reader.Read())
                                            {
                                                string tmp = reader.Name;
                                                if (tmp == "purchaseNumber")
                                                {
                                                    if (reader.ReadString() == lotNumber)
                                                        isWrite = true;
                                                }
                                                if (tmp == "inn")
                                                {
                                                    inn = reader.ReadString();
                                                }
                                                if (tmp == "protocolDate")
                                                {
                                                    protocolDate = Convert.ToDateTime(reader.ReadString());
                                                }
                                                if (tmp == "organizationName")
                                                {
                                                    orgName = reader.ReadString();
                                                    orgName = orgName.Replace("ЗАКРЫТОЕ", "").Replace("АКЦИОНЕРНОЕ", "").Replace("ОБЩЕСТВО", "").Replace("ОТКРЫТОЕ", "")
                                                        .Trim().Replace(" ", "");
                                                    if (organizationName.ContainsKey(orgName + "|" + inn))
                                                    {
                                                        orgName = "";
                                                        inn = "";
                                                    }
                                                    else
                                                    {
                                                        organizationName.Add(orgName + "|" + inn, 0);
                                                    }
                                                }
                                                if (tmp == "criterionCode")
                                                {
                                                    boolOffer = true;
                                                }
                                                if (tmp == "indicatorOffer")
                                                {
                                                    boolOffer = false;
                                                }
                                                if (tmp == "offer" && boolOffer)
                                                {
                                                    string amount = reader.ReadString();
                                                    if (organizationName.ContainsKey(orgName + "|" + inn))
                                                    {
                                                        try
                                                        {
                                                            organizationName[orgName + "|" + inn] += Convert.ToDecimal(amount);
                                                        }
                                                        catch
                                                        {
                                                            organizationName[orgName + "|" + inn] += Convert.ToDecimal(amount.Replace('.', ','));
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        using (XmlReader reader = XmlReader.Create(str))
                                        {
                                            while (reader.Read())
                                            {
                                                string tmp = reader.Name;
                                                if (tmp == "purchaseNumber")
                                                {
                                                    if (reader.ReadString() == lotNumber)
                                                        isWrite = true;
                                                }
                                                if (tmp == "inn")
                                                {
                                                    winners.Add(reader.ReadString());
                                                }
                                                if (tmp == "protocolDate")
                                                {
                                                    protocolDate = Convert.ToDateTime(reader.ReadString());
                                                }
                                            }
                                        }
                                    }
                                    if (isWrite)
                                    {
                                        if (System.IO.File.Exists(pathProtocolsFileLoad + @"\" + item.Name))
                                        {
                                            System.IO.File.Copy(str, pathProtocolsFileLoad + @"\" + un + item.Name);
                                            un++;
                                        }
                                        else
                                            System.IO.File.Copy(str, pathProtocolsFileLoad + @"\" + item.Name);
                                    }
                                }
                            }
                            foreach (var item in dirExtract2.GetFiles())
                            {
                                System.IO.File.Delete(pathProtocolsExtract + @"\" + item.Name);
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    sw.WriteLine(e.ToString());
                }
                sw.Close();
            }
            #endregion

            #region 830
            else if (type == 830)
            {
                string pathZip = @"C:\temp\выгрузка";
                string pathUnZip = @"C:\temp\результат";

                DirectoryInfo dirPathZip = new DirectoryInfo(pathZip);
                DirectoryInfo dirPathUnZip = new DirectoryInfo(pathUnZip);
                int un = 0;
                Directory.SetCurrentDirectory(pathUnZip);
                foreach (var t in dirPathUnZip.GetDirectories())
                {
                    System.IO.Directory.Delete(t.FullName, true);
                }
                try
                {
                    foreach (var items in dirPathZip.GetFiles())
                    {
                        string file = pathZip + @"\" + items.Name;
                        //C:\Temp\7-Zip
                        //ZipFile.ExtractToDirectory(file, pathContractExtract);
                        // Формируем параметры вызова 7z
                        ProcessStartInfo startInfo = new ProcessStartInfo();
                        startInfo.FileName = @"C:\Temp\7-Zip\7z.exe";
                        // Распаковать (для полных путей - x)
                        startInfo.Arguments = " x";
                        // На все отвечать yes
                        startInfo.Arguments += " -y";
                        // Файл, который нужно распаковать
                        startInfo.Arguments += " " + "\"" + file + "\"";
                        // Папка распаковки
                        startInfo.Arguments += " -o" + "\"" + pathUnZip + "\"";
                        startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                        int sevenZipExitCode = 0;
                        using (Process sevenZip = Process.Start(startInfo))
                        {
                            sevenZip.WaitForExit();
                            sevenZipExitCode = sevenZip.ExitCode;
                        }
                        // Если с первого раза не получилось,
                        //пробуем еще раз через 1 секунду
                        if (sevenZipExitCode != 0 && sevenZipExitCode != 1)
                        {
                            using (Process sevenZip = Process.Start(startInfo))
                            {
                                Thread.Sleep(1000);
                                sevenZip.WaitForExit();
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    string temp = "";
                    temp = e.ToString();
                }
                try
                {
                    foreach (var dir in dirPathUnZip.GetDirectories())
                    {
                        foreach (var items in dir.GetFiles())
                        {
                            string file = pathUnZip + @"\" + dir.Name + @"\" + items.Name;
                            //C:\Temp\7-Zip
                            //ZipFile.ExtractToDirectory(file, pathContractExtract);
                            // Формируем параметры вызова 7z
                            ProcessStartInfo startInfo = new ProcessStartInfo();
                            startInfo.FileName = @"C:\Temp\7-Zip\7z.exe";
                            // Распаковать (для полных путей - x)
                            startInfo.Arguments = " x";
                            // На все отвечать yes
                            startInfo.Arguments += " -y";
                            // Файл, который нужно распаковать
                            startInfo.Arguments += " " + "\"" + file + "\"";
                            // Папка распаковки
                            startInfo.Arguments += " -o" + "\"" + pathUnZip + @"\" + dir.Name + "\"";
                            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                            int sevenZipExitCode = 0;
                            using (Process sevenZip = Process.Start(startInfo))
                            {
                                sevenZip.WaitForExit();
                                sevenZipExitCode = sevenZip.ExitCode;
                            }
                            // Если с первого раза не получилось,
                            //пробуем еще раз через 1 секунду
                            if (sevenZipExitCode != 0 && sevenZipExitCode != 1)
                            {
                                using (Process sevenZip = Process.Start(startInfo))
                                {
                                    Thread.Sleep(1000);
                                    sevenZip.WaitForExit();
                                }
                            }
                            System.IO.File.Delete(file);
                        }
                    }
                }
                catch (Exception e)
                {
                    string temp = "";
                    temp = e.ToString();
                }
            }
            #endregion

            #region 82
            else if (type == 82)
            {

                DateTime docPublishDate = new DateTime(9999, 12, 12);
                DateTime docPublishDateTemp = new DateTime();
                //client.ChangeDirectory(TimeoutFTP, "fcs_regions/Samarskaja_obl/contracts");
                string pathNotifikation = @"C:\temp\depstr\notifications\incoming";
                string pathNotifikationExtract = @"C:\temp\depstr\notifications\extract";
                string pathNotifikationFileLoad = @"C:\temp\depstr\notifications\fileLoad";
                DirectoryInfo dirIncoming = new DirectoryInfo(pathNotifikation);
                int un = 0;
                Directory.SetCurrentDirectory(pathNotifikation);


                DirectoryInfo dirExtract = new DirectoryInfo(pathNotifikationExtract);
                foreach (var item in dirExtract.GetFiles())
                {
                    if (item.Name.Contains("Notification"))
                    {
                        string purchaseObjectInfo = "";
                        decimal maxPrice = 0;
                        string code = "";
                        bool isWrite = false;

                        string str = pathNotifikationExtract + @"\" + item.Name;
                        using (XmlReader reader = XmlReader.Create(str))
                        {
                            while (reader.Read())
                            {
                                string tmp = reader.Name;
                                if (tmp == "purchaseObjectInfo")
                                {
                                    purchaseObjectInfo = reader.ReadString();
                                }
                                if (tmp == "maxPrice")
                                {
                                    string price = reader.ReadString();
                                    Console.WriteLine(price);
                                    try
                                    {
                                        maxPrice = Convert.ToDecimal(price);
                                    }
                                    catch
                                    {
                                        try
                                        {
                                            maxPrice = Convert.ToDecimal(price.Replace('.', ','));
                                        }
                                        catch
                                        {
                                        }
                                    }
                                }
                                if (tmp == "code")
                                {
                                    code = reader.ReadString();
                                    Console.WriteLine(code);

                                }
                                if (tmp == "purchaseNumber")
                                {
                                    isWrite = true;
                                }
                                if (tmp == "docPublishDate")
                                {
                                    docPublishDateTemp = Convert.ToDateTime(reader.ReadString());
                                }
                            }
                        }
                        /*if (isWrite)
                        {
                            if (System.IO.File.Exists(pathNotifikationFileLoad + @"\" + item.Name))
                            {
                                System.IO.File.Copy(str, pathNotifikationFileLoad + @"\" + un + item.Name);
                                un++;
                            }
                            else
                                System.IO.File.Copy(str, pathNotifikationFileLoad + @"\" + item.Name);
                        }*/

                    }
                }
            }
            #endregion

            //согрузка ПС КАРТОТЕКА
            #region 91
            else if (type == 91)
            {
                StreamWriter sw = new StreamWriter(@"C:\temp\depstr\error.log", false);
                string lotNumber = "";
                FtpClient client = new FtpClient();
                //Задаём параметры клиента.
                client.PassiveMode = true; //Включаем пассивный режим.
                int TimeoutFTP = 30000; //Таймаут.
                string FTP_SERVER = "ftp.zakupki.gov.ru";
                //Подключаемся к FTP серверу.
                client.Connect(TimeoutFTP, FTP_SERVER, 21);
                client.Login(TimeoutFTP, "free", "free");
                string pathContracts = @"C:\temp\depstr\contracts\incoming";
                string pathContractsExtract = @"C:\temp\depstr\contracts\extract";
                string pathContractsFileLoad = @"C:\temp\depstr\contracts\fileLoad";
                DirectoryInfo dirIncoming3 = new DirectoryInfo(pathContracts);
                int un = 0;
                
                client.ChangeDirectory(TimeoutFTP, "fcs_regions/Samarskaja_obl/contracts");
                Directory.SetCurrentDirectory(pathContracts);
                foreach (var t in client.GetDirectoryList(TimeoutFTP))
                {
                    if (t.Name.Substring(t.Name.Length - 3) == "zip" && (t.Name.Contains("2015") || t.Name.Contains("2014"))
                            && !System.IO.File.Exists(Directory.GetCurrentDirectory() + @"\" + t.Name))
                    {
                        string file = Directory.GetCurrentDirectory() + @"\" + t.Name;
                        try
                        {
                            client.GetFile(TimeoutFTP, file, t.Name);
                        }
                        catch
                        {
                            System.IO.File.Delete(file);
                        }
                    }
                }
                client.Disconnect(TimeoutFTP);

                DirectoryInfo dirExtract3 = new DirectoryInfo(pathContractsExtract);
                foreach (var item in dirExtract3.GetFiles())
                {
                    System.IO.File.Delete(pathContractsExtract + @"\" + item.Name);
                }
                try
                {
                    for (int i = 0; i < 1; i++)
                    {
                        lotNumber = "0142300024514000027";
                        foreach (var items in dirIncoming3.GetFiles())
                        {

                            //sw.WriteLine("file= " + pathNotifikation + @"\" + items.Name);
                            string file = pathContracts + @"\" + items.Name;
                            //C:\Temp\7-Zip
                            //ZipFile.ExtractToDirectory(file, pathContractExtract);
                            // Формируем параметры вызова 7z
                            ProcessStartInfo startInfo = new ProcessStartInfo();
                            startInfo.FileName = @"C:\Temp\7-Zip\7z.exe";
                            // Распаковать (для полных путей - x)
                            startInfo.Arguments = " e";
                            // На все отвечать yes
                            startInfo.Arguments += " -y";
                            // Файл, который нужно распаковать
                            startInfo.Arguments += " " + "\"" + file + "\"";
                            // Папка распаковки
                            startInfo.Arguments += " -o" + "\"" + pathContractsExtract + "\"";
                            startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                            int sevenZipExitCode = 0;
                            using (Process sevenZip = Process.Start(startInfo))
                            {
                                sevenZip.WaitForExit();
                                sevenZipExitCode = sevenZip.ExitCode;
                            }
                            // Если с первого раза не получилось,
                            //пробуем еще раз через 1 секунду
                            if (sevenZipExitCode != 0 && sevenZipExitCode != 1)
                            {
                                using (Process sevenZip = Process.Start(startInfo))
                                {
                                    Thread.Sleep(1000);
                                    sevenZip.WaitForExit();
                                }
                            }

                            foreach (var item in dirExtract3.GetFiles())
                            {
                                try
                                {
                                    if (!item.Name.Contains("Notificati;liujluijlion"))
                                    {
                                        bool isWrite = false;
                                        string str = pathContractsExtract + @"\" + item.Name;
                                        if (item.Name.Contains("0142300024514000027"))
                                            isWrite = true;


                                        if (isWrite)
                                        {
                                            if (System.IO.File.Exists(pathContractsFileLoad + @"\" + item.Name))
                                            {
                                                System.IO.File.Copy(str, pathContractsFileLoad + @"\" + un + item.Name);
                                                un++;
                                            }
                                            else
                                                System.IO.File.Copy(str, pathContractsFileLoad + @"\" + item.Name);
                                        }
                                    }
                                }
                                catch (Exception e)
                                {
                                    sw.WriteLine(e.ToString());
                                    continue;
                                }
                            }
                            foreach (var item in dirExtract3.GetFiles())
                            {
                                System.IO.File.Delete(pathContractsExtract + @"\" + item.Name);
                            }

                        }
                        
                    }
                }
                catch (Exception e)
                {
                    sw.WriteLine(e.ToString());
                }
                sw.Close();
            }
            #endregion

            //согрузка ПС КАРТОТЕКА
            #region 90
            else if (type == 90)
            {
                pg.UpdateLs();
            }
            #endregion

            #region 81 meta_attr_id
            else if (type == 81)
            {
                StreamWriter sw1 = new StreamWriter(@"C:\temp\meta_atr.txt", false);
                string fileName = @"C:\temp\meta.txt";
                string[] stringSeparators = new string[] { "meta_attribute_id" };
                string[] allText = File.ReadAllLines(fileName);         //чтение всех строк файла в массив строк
                List<string> attr = new List<string>();
                //string str = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Split(stringSeparators, StringSplitOptions.None)[0].Replace(" ", "");
                foreach (string s in allText)
                {
                    if (s.Contains("meta_attribute_id"))
                    {
                        if (s.Split(stringSeparators, StringSplitOptions.None)[1].Replace('=', ' ').Trim().Length >= 8
                            && !attr.Contains(s.Split(stringSeparators, StringSplitOptions.None)[1].Replace('=', ' ').Trim().Substring(0, 7)))
                            attr.Add(s.Split(stringSeparators, StringSplitOptions.None)[1].Replace('=', ' ').Trim().Substring(0, 7));
                    }
                }
                foreach (string s in attr)
                {
                    sw1.WriteLine("metaAtrId.Add(\"" + s + "\");");
                }
                sw1.Close();
            }
            #endregion


            #region 92 Begin
            else if (type == 92)
            {
                DataTable obj = ora.SelectInfo(10);
                var wb = new XLWorkbook(@"C:\temp\reportOrg.xlsx");
                int rowMove = 1;
                for (int i = 0; i < obj.Rows.Count; i++)
                {
                    string contakts = "";
                    string address = obj.Rows[i][2].ToString();
                    if (obj.Rows[i][2].ToString() != obj.Rows[i][7].ToString() && obj.Rows[i][7] != null && obj.Rows[i][7].ToString() != "")
                        address += "/" + obj.Rows[i][7].ToString();
                    if (obj.Rows[i][3] != null && obj.Rows[i][3].ToString() != "")
                        contakts += obj.Rows[i][3].ToString() + ",";
                    if (obj.Rows[i][4] != null && obj.Rows[i][4].ToString() != "")
                        contakts += obj.Rows[i][4].ToString() + ",";
                    if (obj.Rows[i][5] != null && obj.Rows[i][5].ToString() != "")
                        contakts += obj.Rows[i][5].ToString() + ",";
                    if (contakts.Length > 1)
                        contakts = contakts.Substring(0, contakts.Length - 1);
                    rowMove++;
                    wb.Worksheet(1).Row(rowMove).Cell(1).Value = obj.Rows[i][1].ToString();
                    wb.Worksheet(1).Row(rowMove).Cell(2).Value = address;
                    wb.Worksheet(1).Row(rowMove).Cell(3).Value = contakts;
                    wb.Worksheet(1).Row(rowMove).Cell(4).Value = obj.Rows[i][6].ToString();
                }
                //DataTable obj2 = ora.SelectInfo(20);
                //for (int i = 0; i < obj2.Rows.Count; i++)
                //{
                //    string contakts = "";
                //    string address = obj2.Rows[i][2].ToString();
                //    if (obj2.Rows[i][2].ToString() != obj2.Rows[i][7].ToString() && obj2.Rows[i][7] != null && obj2.Rows[i][7].ToString() != "")
                //        address += "/" + obj2.Rows[i][7].ToString();
                //    if (obj2.Rows[i][3] != null && obj2.Rows[i][3].ToString() != "")
                //        contakts += obj2.Rows[i][3].ToString() + ",";
                //    if (obj2.Rows[i][4] != null && obj2.Rows[i][4].ToString() != "")
                //        contakts += obj2.Rows[i][4].ToString() + ",";
                //    if (obj2.Rows[i][5] != null && obj2.Rows[i][5].ToString() != "")
                //        contakts += obj2.Rows[i][5].ToString() + ",";
                //    if (contakts.Length > 1)
                //        contakts = contakts.Substring(0, contakts.Length - 1);
                //    rowMove++;
                //    wb.Worksheet(2).Row(rowMove).Cell(1).Value = obj2.Rows[i][1].ToString();
                //    wb.Worksheet(2).Row(rowMove).Cell(2).Value = address;
                //    wb.Worksheet(2).Row(rowMove).Cell(3).Value = contakts;
                //    wb.Worksheet(2).Row(rowMove).Cell(4).Value = obj2.Rows[i][6].ToString();
                //}
                //DataTable obj3 = ora.SelectInfo();
                //for (int i = 0; i < obj3.Rows.Count; i++)
                //{
                //    string contakts = "";
                //    string address = obj3.Rows[i][2].ToString();
                //    if (obj3.Rows[i][2].ToString() != obj3.Rows[i][7].ToString() && obj3.Rows[i][7] != null && obj3.Rows[i][7].ToString() != "")
                //        address += "/" + obj3.Rows[i][7].ToString();
                //    if (obj3.Rows[i][3] != null && obj3.Rows[i][3].ToString() != "")
                //        contakts += obj3.Rows[i][3].ToString() + ",";
                //    if (obj3.Rows[i][4] != null && obj3.Rows[i][4].ToString() != "")
                //        contakts += obj3.Rows[i][4].ToString() + ",";
                //    if (obj3.Rows[i][5] != null && obj3.Rows[i][5].ToString() != "")
                //        contakts += obj3.Rows[i][5].ToString() + ",";
                //    if (contakts.Length > 0)
                //        contakts = contakts.Substring(0, contakts.Length - 1);
                //    rowMove++;
                //    wb.Worksheet(3).Row(rowMove).Cell(1).Value = obj3.Rows[i][1].ToString();
                //    wb.Worksheet(3).Row(rowMove).Cell(2).Value = address;
                //    wb.Worksheet(3).Row(rowMove).Cell(3).Value = contakts;
                //    wb.Worksheet(3).Row(rowMove).Cell(4).Value = obj3.Rows[i][6].ToString();
                //}
                wb.Save();
            }
            #endregion

            //перекидка
            #region 93
            else if (type == 93)
            {
                var book = new XLWorkbook(@"C:\temp\SpLs_93684578133.xlsx");
                for (int i = 4; i <= 136; i++)
                {
                    pg.InsertRoomCount(book.Worksheet(1).Row(i).Cell(1).Value.ToString(), book.Worksheet(1).Row(i).Cell(2).Value.ToString());
                }
                book.Save();
            }
            #endregion

            //проставка pkod
            #region 94
            else if (type == 94)
            {
                var book = new XLWorkbook(@"C:\temp\Реестр для Димы.xlsx");
                for (int i = 2; i <= 293; i++)
                {
                    string pkod = pg.SelectPkod(Convert.ToString(book.Worksheet(1).Row(i).Cell(6).Value), Convert.ToString(book.Worksheet(1).Row(i).Cell(7).Value));
                    if (pkod != "00" && pkod != "0")
                    {
                        book.Worksheet(1).Row(i).Cell(1).Value = pkod;
                    }
                    else
                    {
                        book.Worksheet(1).Row(i).Cell(1).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                }
                book.Save();
            }
            #endregion

            //обновление даты оплат
            #region 95
            else if (type == 95)
            {
                var book = new XLWorkbook(@"C:\temp\Реестр для Димы.xlsx");
                for (int i = 2; i <= 293; i++)
                {
                    int count = pg.UpdateDatePaid(Convert.ToString(book.Worksheet(1).Row(i).Cell(1).Value), Convert.ToString(book.Worksheet(1).Row(i).Cell(2).Value),
                        Convert.ToString(book.Worksheet(1).Row(i).Cell(3).Value), Convert.ToString(book.Worksheet(1).Row(i).Cell(4).Value),
                        Convert.ToString(book.Worksheet(1).Row(i).Cell(5).Value));
                    if (count == 0)
                    {
                         book.Worksheet(1).Row(i).Cell(1).Style.Fill.BackgroundColor = XLColor.Yellow;
                         book.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Yellow;
                         book.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Yellow;
                         book.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Yellow;
                         book.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Yellow;
                         book.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Yellow;
                         book.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                    else if (count == 2)
                    {
                         book.Worksheet(1).Row(i).Cell(1).Style.Fill.BackgroundColor = XLColor.Green;
                         book.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Green;
                         book.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Green;
                         book.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Green;
                         book.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Green;
                         book.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Green;
                         book.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Green;
                    }
                }
                book.Save();
            }
            #endregion

            //Собственники с пробегом по всем файлам
            #region 96
            else if (type == 96)
            {
                DirectoryInfo dir = new DirectoryInfo(@"C:\temp\houses5");
                string[] stringSeparators = new string[] { "Кв." };
                foreach (var item in dir.GetFiles())
                {
                    Console.WriteLine(item.Name);
                    var wb2 = new XLWorkbook(@"C:\temp\houses5\" + item.Name);
                    for (int i = 10; i <= 1000; i++)
                    {
                        if (wb2.Worksheet(1).Row(i).Cell(3).Value == null || Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim() == "")
                            break;
                        try
                        {
                            //string str = ora.UpdatePeople(item.Name.Substring(0, item.Name.Length - 5),
                            //    Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim().Replace('.', ' ').Trim(),
                            //    Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim(),
                            //    Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value),
                            //    Convert.ToString(wb2.Worksheet(1).Row(i).Cell(7).Value),
                            //    Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value));
                            //if (str != "ЗАГРУЖЕНО")
                            //    Console.WriteLine(str);
                        }
                        catch
                        {

                        }

                    }
                    wb2.Save();
                }
            }
            #endregion

            //InsertNewLS
            #region 97
            else if (type == 97)
            {
                int flat_num = 2;
                for (int i = 14568393; i <= 14568393+95; i++)
                {
                    pg.InsertNewLS(i, flat_num);
                    flat_num++;
                }
            }
            #endregion

            #region 98
            else if (type == 98)
            {
                DataTable dtHouse = new DataTable();
                DataTable dtPeople = new DataTable();
                var wb2 = new XLWorkbook(@"C:\temp\Сведения по нежил.помещениям.xlsx");
                for (int i = 5; i <= 5; i++)
                {
                    string str = ora.InsertOffice(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(7).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(8).Value).Trim());

                    if (str != "ЗАГРУЖЕНО")
                        Console.WriteLine(str);                
                }
                wb2.Save();
            }
            #endregion

            //Загрузка лс в биллнг с параметрами
            #region 99
            else if (type == 99)
            {
                var wb2 = new XLWorkbook(@"C:\temp\с.п.Тимофеевка (2).xlsx");
                
                for (int i = 14; i <= 78; i++)
                {
                    Int32 nzp_kvar = pg.InsertFio(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim());

                    Int32 res = pg.InsertPrmByKvar(nzp_kvar,
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim().Replace(',','.'), 5);
                    if (res == 0)
                    {
                        wb2.Worksheet(1).Row(i).Cell(1).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }

                    res = pg.InsertPrmByKvar(nzp_kvar,
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim().Replace(',', '.'), 4);
                    if (res == 0)
                    {
                        wb2.Worksheet(1).Row(i).Cell(1).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                }

                wb2.Save();
            }
            #endregion

            //Загрузка лс в биллнг с параметрами
            #region 100
            else if (type == 100)
            {
                var wb2 = new XLWorkbook(@"C:\temp\aukExp.xlsx");
                for (int i = 1; i <= 208; i++)
                {
                    pg.InsertFio(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim(),
                        "",
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(7).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim());
                }
                wb2.Save();
            }
            #endregion

            //Загрузка лс в биллнг с параметрами
            #region 101
            else if (type == 101)
            {
                var wb2 = new XLWorkbook(@"C:\temp\Tash188.xlsx");
                for (int i = 1; i <= 208; i++)
                {
                    pg.InsertRoomCount(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(45).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(13).Value).Trim() + " " +
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(14).Value).Trim() + " " +
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(15).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(8).Value).Trim());
                }
                wb2.Save();
            }
            #endregion

            //Загрузка параметров
            #region 102
            else if (type == 102)
            {
                var wb2 = new XLWorkbook(@"C:\temp\aukExp.xlsx");
                for (int i = 1; i <= 208; i++)
                {
                    if (i == 25 || i == 184)
                    {
                        int t = pg.InsertPrm(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim(),
                       Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim(),
                       Convert.ToString(wb2.Worksheet(1).Row(i).Cell(7).Value).Trim(),
                       Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim(),
                       Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim(),
                       Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value).Trim(),
                       Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value).Trim());
                        if (t == 0)
                        {
                            wb2.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                        }
                    }
                   
                }
                wb2.Save();
            }
            #endregion

            //Загрузка параметров
            #region 1020
            else if (type == 1020)
            {
                var wb2 = new XLWorkbook(@"C:\temp\ЖЭУ 1 для загрузки Люда ПРОВЕРЕН.xlsx");
                Console.Write("Введите имя БД:");
                string database = Console.ReadLine();
                for (int i = 2; i <= 93; i++)
                {
                    string kvar = pg.SelectNzpKvarByKvarDom(database, wb2.Worksheet(1).Row(i).Cell(1).Value.ToString(), 7155108);
                    string nzp_kvar = "";
                    if (kvar.Split('|')[1] == "Найдено")
                    {
                        nzp_kvar = kvar.Split('|')[0];
                    }
                    else
                    {
                        wb2.Worksheet(1).Row(i).Cell(1).Style.Fill.BackgroundColor = XLColor.Yellow;
                        continue;
                    }
               
                    //pg.UpdateFio(database, nzp_kvar,
                    //    Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim().ToUpper());

                    DataTable t = pg.TestFio(database, nzp_kvar);
                    if (t.Rows.Count != 0)
                    {
                        wb2.Worksheet(1).Row(i).Cell(1).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                  
                    //int t = pg.InsertPrm(database, nzp_kvar,
                    //    Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim(), 4);
                    //if (t == 0)
                    //{
                    //    wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Yellow;
                    //}
                    //t = pg.InsertPrm(database, nzp_kvar,
                    //    Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim(), 1010270);
                    //if (t == 0)
                    //{
                    //    wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Yellow;
                    //}
                    //t = pg.InsertPrm(database, nzp_kvar,
                    //    Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value).Trim(), 5);
                    //if (t == 0)
                    //{
                    //    wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Yellow;
                    //}
                }
                wb2.Save();
            }
            #endregion

            //Загрузка лс в биллнг с параметрами
            #region 103
            else if (type == 103)
            {
                var wb2 = new XLWorkbook(@"C:\temp\Pkod.xlsx");
                for (int i = 1; i <= 3498; i++)
                {
                    int t = pg.UpdatePkod(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim(),
                        Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim());
                    if (t == 0)
                    {
                        wb2.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                }
                wb2.Save();
            }
            #endregion

            //Загрузка лс в биллнг с параметрами
            #region 104
            else if (type == 104)
            {
                var wb2 = new XLWorkbook(@"C:\temp\Копия Действующие УО ТСЖ ЖСК.xlsx");
                for (int i = 3; i <= 333; i++)
                {
                    string fio = ora.SelectFio(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value).Trim(), 293833);
                    wb2.Worksheet(1).Row(i).Cell(9).Value = fio;
                    fio = ora.SelectFio(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value).Trim(), 293834);
                    wb2.Worksheet(1).Row(i).Cell(10).Value = fio;
                    fio = ora.SelectFio(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value).Trim(), 293837);
                    wb2.Worksheet(1).Row(i).Cell(11).Value = fio;
                    fio = ora.SelectFio(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value).Trim(), 293839);
                    wb2.Worksheet(1).Row(i).Cell(12).Value = fio;
                    fio = ora.SelectFio(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value).Trim(), 293864);
                    wb2.Worksheet(1).Row(i).Cell(13).Value = fio;
                    fio = ora.SelectFio(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value).Trim(), 293865);
                    wb2.Worksheet(1).Row(i).Cell(14).Value = fio;
                    fio = ora.SelectFio(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value).Trim(), 293867);
                    wb2.Worksheet(1).Row(i).Cell(15).Value = fio;
                }
                wb2.Save();
            }
            #endregion

            //создание dbf 
            #region 105
            else if (type == 105)
            {
                var wb2 = new XLWorkbook(@"C:\temp\ИПУ(2).xlsx");
                string dat_uchet = Convert.ToString(wb2.Worksheet(1).Row(3).Cell(6).Value).Trim();
                DateTime date_pay = Convert.ToDateTime(wb2.Worksheet(1).Row(3).Cell(7).Value);
                for (int i = 5; i <= 5; i++)
                {
                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim() != "")
                    {
                        List<string> nzp_kvar = pg.SelectNzpKvar(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim(),
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim());
                        if (nzp_kvar != null)
                        {
                            string addCounters = pg.AddCounter(
                                nzp_kvar[0],
                                nzp_kvar[1],
                                Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim(),
                                dat_uchet,
                                date_pay,
                                Convert.ToDecimal(wb2.Worksheet(1).Row(i).Cell(6).Value),
                                Convert.ToDecimal(wb2.Worksheet(1).Row(i).Cell(7).Value)
                            );
                            if (addCounters == "Success")
                            {

                            }
                            else
                            {
                                
                            }
                        }
                    }
                    else
                    {
                        break;
                    }
                }
                wb2.Save();
            }
            #endregion

            //Загрузка счетчиком для сайта
            #region 106
            else if (type == 106)
            {
                Console.Write("Введите имя БД:");
                string database = Console.ReadLine();
                DataTable dt = pg.SelectCounters(database);
                XmlTextWriter myXml = new XmlTextWriter(@"C:\Temp\" + database + ".xml", System.Text.Encoding.Default);
                StreamWriter errorRow = new StreamWriter(@"C:\Temp\error3.txt", false, Encoding.Default);
                myXml.Formatting = Formatting.None;
                myXml.WriteStartDocument(true);
                myXml.WriteStartElement("Фаил");
                string num_ls = "";
                string serv = "";
                string counter = "";
                string counter_num = "";
                try
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (num_ls != dt.Rows[i][2].ToString().Trim())
                        {
                            if (i != 0)
                            {
                                myXml.WriteEndElement();
                                myXml.WriteEndElement();
                                myXml.WriteEndElement();
                                myXml.WriteEndElement();
                                myXml.WriteEndElement();
                                myXml.WriteEndElement();
                            }
                            num_ls = dt.Rows[i][2].ToString().Trim();
                            serv = dt.Rows[i][3].ToString().Trim();
                            counter = dt.Rows[i][4].ToString().Trim();
                            counter_num = dt.Rows[i][5].ToString().Trim();
                            string cnt_num = "";
                            int j = i;
                            while (dt.Rows[j][4].ToString() == counter)
                            {
                                if (dt.Rows[j][5].ToString() != "")
                                    cnt_num = dt.Rows[j][5].ToString();
                                j++;
                                if (j == dt.Rows.Count)
                                    break;
                            }
                            myXml.WriteStartElement("ЛС");
                            myXml.WriteElementString("Адрес", dt.Rows[i][0].ToString().Trim());
                            myXml.WriteElementString("ФИО", dt.Rows[i][1].ToString().Trim());
                            myXml.WriteElementString("Номер_ЛС", dt.Rows[i][2].ToString().Trim());
                            myXml.WriteStartElement("Услуги");
                            myXml.WriteStartElement("Услуга");
                            myXml.WriteElementString("Наименование", dt.Rows[i][3].ToString().Trim());
                            myXml.WriteStartElement("Счетчики");
                            myXml.WriteStartElement("Счетчик");
                            myXml.WriteElementString("ID", dt.Rows[i][4].ToString().Trim());
                            myXml.WriteElementString("Номер", cnt_num);
                            myXml.WriteStartElement("Показания");
                            myXml.WriteStartElement("Показание");
                            myXml.WriteElementString("Год", dt.Rows[i][8].ToString().Trim());
                            myXml.WriteElementString("Месяц", dt.Rows[i][6].ToString().Trim());
                            myXml.WriteElementString("Значение", dt.Rows[i][7].ToString().Trim());
                            myXml.WriteEndElement();
                        }
                        else
                        {
                            if (serv != dt.Rows[i][3].ToString().Trim())
                            {
                                serv = dt.Rows[i][3].ToString().Trim();
                                counter = dt.Rows[i][4].ToString().Trim();
                                counter_num = dt.Rows[i][5].ToString().Trim();
                                string cnt_num = "";
                                int j = i;
                                while (dt.Rows[j][4].ToString() == counter)
                                {
                                    if (dt.Rows[j][5].ToString() != "")
                                        cnt_num = dt.Rows[j][5].ToString();
                                    j++;
                                    if (j == dt.Rows.Count)
                                        break;
                                }
                                myXml.WriteEndElement();
                                myXml.WriteEndElement();
                                myXml.WriteEndElement();
                                myXml.WriteEndElement();
                                myXml.WriteStartElement("Услуга");
                                myXml.WriteElementString("Наименование", dt.Rows[i][3].ToString().Trim());
                                myXml.WriteStartElement("Счетчики");
                                myXml.WriteStartElement("Счетчик");
                                myXml.WriteElementString("ID", dt.Rows[i][4].ToString().Trim());
                                myXml.WriteElementString("Номер", cnt_num);
                                myXml.WriteStartElement("Показания");
                                myXml.WriteStartElement("Показание");
                                myXml.WriteElementString("Год", dt.Rows[i][8].ToString().Trim());
                                myXml.WriteElementString("Месяц", dt.Rows[i][6].ToString().Trim());
                                myXml.WriteElementString("Значение", dt.Rows[i][7].ToString().Trim());
                                myXml.WriteEndElement();
                            }
                            else
                            {
                                if (counter != dt.Rows[i][4].ToString().Trim())
                                {                                                             
                                    counter = dt.Rows[i][4].ToString().Trim();
                                    string cnt_num = "";
                                    int j = i;   
                                    while (dt.Rows[j][4].ToString() == counter)
                                    {
                                        if (dt.Rows[j][5].ToString() != "")
                                            cnt_num = dt.Rows[j][5].ToString();
                                        j++;
                                        if (j == dt.Rows.Count)
                                            break;
                                    }
                                    myXml.WriteEndElement();
                                    myXml.WriteEndElement();
                                    myXml.WriteStartElement("Счетчик");
                                    myXml.WriteElementString("ID", dt.Rows[i][4].ToString().Trim());
                                    myXml.WriteElementString("Номер", cnt_num);
                                    myXml.WriteStartElement("Показания");
                                    myXml.WriteStartElement("Показание");
                                    myXml.WriteElementString("Год", dt.Rows[i][8].ToString().Trim());
                                    myXml.WriteElementString("Месяц", dt.Rows[i][6].ToString().Trim());
                                    myXml.WriteElementString("Значение", dt.Rows[i][7].ToString().Trim());
                                    myXml.WriteEndElement();
                                }
                                else
                                {
                                    myXml.WriteStartElement("Показание");
                                    myXml.WriteElementString("Год", dt.Rows[i][8].ToString().Trim());
                                    myXml.WriteElementString("Месяц", dt.Rows[i][6].ToString().Trim());
                                    myXml.WriteElementString("Значение", dt.Rows[i][7].ToString().Trim());
                                    myXml.WriteEndElement();
                                }
                            }

                        }

                    }
                    myXml.WriteEndElement();
                    myXml.WriteEndElement();
                    myXml.WriteEndElement();
                    myXml.WriteEndElement();
                    myXml.WriteEndElement();
                    myXml.WriteEndElement();
                }
                catch (Exception e)
                {
                    errorRow.WriteLine(e.ToString());
                    errorRow.Close();
                }
               
                myXml.WriteEndElement();
                myXml.Flush();
                myXml.Close();
            }
            #endregion

            //Импорт квартир в РЦ
            #region 107
            else if (type == 107)
            {
                string[] separator = new string[] { "кв." };
                string[] separator1 = new string[] { "неж." };
                string[] separator2 = new string[] { ", ком." };
                var wb2 = new XLWorkbook(@"C:\temp\Революционная 50.xlsx");
                for (int i = 279; i <= 285; i++)
                {
                    string num_ls = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim().Substring(5);
                    string fio = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim();
                    string nkvar = "";
                    Int32 ikvar = 0;
                    string nkvar_n = "";
                    string str = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim();
                    if (!Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim().Contains(", ком."))
                    {
                        if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim().Contains("неж."))
                        {
                            nkvar = "неж. " + Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value)
                                .Split(separator1, StringSplitOptions.None)[1].Trim();
                            if (!Int32.TryParse(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value)
                                .Split(separator1, StringSplitOptions.None)[1].Trim(), out ikvar))
                            {
                                ikvar = 0;
                            }
                        }
                        else
                        {
                            nkvar = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value)
                                .Split(separator, StringSplitOptions.None)[1].Trim();
                            if (!Int32.TryParse(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value)
                                .Split(separator, StringSplitOptions.None)[1].Trim(), out ikvar))
                            {
                                ikvar = 0;
                            }
                        }
                            
                    }
                    else
                    {
                        nkvar =
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value)
                                .Split(separator, StringSplitOptions.None)[1].Trim()
                                    .Split(separator2, StringSplitOptions.None)[0].Trim();
                        nkvar_n = 
                            Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value)
                                .Split(separator, StringSplitOptions.None)[1].Trim()
                                    .Split(separator2, StringSplitOptions.None)[1].Trim();
                        if (!Int32.TryParse(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value)
                                .Split(separator, StringSplitOptions.None)[1].Trim()
                                    .Split(separator2, StringSplitOptions.None)[0].Trim(), out ikvar))
                        {
                            ikvar = 0;
                        }
                    }
                    string res = pg.InsertKvar(num_ls, fio, nkvar, ikvar, nkvar_n);
                    if(res != "Success")
                        Console.WriteLine(res + "|||" + nkvar);
                }
                wb2.Save();
            }
            #endregion

            //Из Excel формируем список параметров для записи в БД Билилнг
            #region 108
            else if (type == 108)
            {
                var wb2 = new XLWorkbook(@"C:\Users\WCSMR-HP\Desktop\Импорт\ул_Революционная_ д_44.xlsx");
                var wb = new XLWorkbook(@"C:\temp\ImportPrm.xlsx");
                int row = 532;
                for (int i = 24; i <= 20869; i++)
                {
                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim().Contains("40005")
                        && Convert.ToString(wb2.Worksheet(1).Row(i+1).Cell(1).Value).Trim() == "")
                    {
                        wb.Worksheet(1).Row(row).Cell(1).Value = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim();
                        while (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim() != "плательщик")
                        {
                            i++;
                        }
                        wb.Worksheet(1).Row(row).Cell(2).Value = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim();
                        wb.Worksheet(1).Row(row).Cell(3).Value = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value).Trim();
                        wb.Worksheet(1).Row(row).Cell(4).Value = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value).Trim().Split('/')[0];
                        row++;
                    }
                }
                wb.Save();
                wb2.Save();
            }
            #endregion

            //Из Excel формируем список параметров для записи в БД Билилнг
            #region 109
            else if (type == 109)
            {
                var wb2 = new XLWorkbook(@"C:\Temp\120.xlsx");
                string address = "";
                Int32 nzp_serv;
                Int32 nzp_cnt;
                string[] separator = new string[] { "кв." };
                string[] separator1 = new string[] { "неж." };
                Console.Write("Введите наименование БД:");
                string database = Console.ReadLine();
                List<string> kvarParams = new List<string>();
                for (int i = 5; i <= 1281; i++)
                {
                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim() != address
                        && Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim() != "")
                    {
                        address = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim();
                        if (address.Contains("нежи."))
                        {
                            kvarParams = pg.SelectKvarParams("",database, address.Split(separator1, StringSplitOptions.None)[1].Trim(), 7155107);
                        }

                        else
                        {
                            kvarParams = pg.SelectKvarParams("", database, address.Split(separator, StringSplitOptions.None)[1].Trim(), 7155107);
                        }
                        
                    }
                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim() == "")
                    {
                        wb2.Worksheet(1).Row(i).Cell(2).Value = address;
                    }
                    if (kvarParams == null)
                    {
                        wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                    else
                    {
                        String serv = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim();
                        switch (serv)
                        {
                            case "ГВС":
                            {
                                nzp_serv = 9;
                                nzp_cnt = 3;
                                break;
                            }
                            case  "ХВС":
                            {
                                nzp_serv = 6;
                                nzp_cnt = 2;
                                break;
                            }
                            case "ЭЛ":
                            case "ЭЛД":
                            {
                                nzp_serv = 25;
                                nzp_cnt = 1;
                                break;
                            }
                            case "ЭЛН":
                            {
                                nzp_serv = 210;
                                nzp_cnt = 7;
                                break;
                            }
                            default:
                            {
                                nzp_serv = 0;
                                nzp_cnt = 0;
                                break;
                            }
                        }
                        if (nzp_serv == 0)
                        {
                            wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Yellow;
                            wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Yellow;
                            wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Yellow;
                            wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Yellow;
                            wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Yellow;
                            wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Yellow;
                        }
                        else
                        {
                            int nzp_counter = pg.SelectNzpCounter("", database, kvarParams[0], nzp_serv, Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim(), "", "");
                            if(nzp_counter == 0)
                                nzp_counter = pg.InsertCounter(kvarParams[0], nzp_serv, Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim(), nzp_cnt, "");
                            int t = pg.InsertCounterVal("", database, kvarParams[0], kvarParams[1], nzp_serv, Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim(),
                                "01.10.2015", Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value).Trim(), nzp_counter);
                            if (t == 0)
                            {
                                wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Yellow;
                                wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Yellow;
                                wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Yellow;
                                wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Yellow;
                                wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Yellow;
                                wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Yellow;
                            }
                            //t = pg.InsertCounterVal(kvarParams[0], kvarParams[1], nzp_serv, Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim(),
                            //    "01.09.2015", Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value).Trim(), nzp_counter);
                            //if (t == 0)
                            //{
                            //    wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Yellow;
                            //    wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Yellow;
                            //    wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Yellow;
                            //    wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Yellow;
                            //    wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Yellow;
                            //    wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Yellow;
                            //}
                            //t = pg.InsertCounterVal(kvarParams[0], kvarParams[1], nzp_serv, Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim(),
                            //    "01.06.2015", Convert.ToString(wb2.Worksheet(1).Row(i).Cell(7).Value).Trim(), nzp_counter);
                            //if (t == 0)
                            //{
                            //    wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Yellow;
                            //    wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Yellow;
                            //    wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Yellow;
                            //    wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Yellow;
                            //    wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Yellow;
                            //    wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Yellow;
                            //}
                        }
                    }
                }
                wb2.Save();
            }
            #endregion

            #region 1091
            else if (type == 1091)
            {
                var wb2 = new XLWorkbook(@"C:\Temp\template_ipu0.xlsx");
                string address = "";
                Int32 nzp_serv;
                Int32 nzp_cnt;
                Int32 end;
                string[] separator = new string[] { "кв." };
                string[] separator1 = new string[] { "неж." };
                Console.Write("Введите наименование БД:");
                string database = Console.ReadLine();
                Console.Write("Введите конечную строку:");
                end = Convert.ToInt32(Console.ReadLine());
                List<string> kvarParams = new List<string>();
                for (int i = 5; i <= end; i++)
                {
                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim() != address
                        && Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim() != "")
                    {
                        address = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim();
                        if (address.Contains("нежи."))
                        {
                            kvarParams = pg.SelectKvarParams("", database, address.Split(separator1, StringSplitOptions.None)[1].Trim(), 7155106);
                        }

                        else
                        {
                            kvarParams = pg.SelectKvarParams("", database, address.Split(separator, StringSplitOptions.None)[1].Trim(), 7155106);
                        }
                    }
                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim() == "")
                    {
                        wb2.Worksheet(1).Row(i).Cell(2).Value = kvarParams[0];
                    }
                    if (kvarParams == null)
                    {
                        wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                    else
                    {
                        wb2.Worksheet(1).Row(i).Cell(2).Value = kvarParams[0];
                        String serv = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim();
                        switch (serv)
                        {
                            case "ГВС":
                                {
                                    nzp_serv = 9;
                                    nzp_cnt = 3;
                                    break;
                                }
                            case "ХВС":
                                {
                                    nzp_serv = 6;
                                    nzp_cnt = 2;
                                    break;
                                }
                            case "ЭЛ":
                            case "ЭЛД":
                                {
                                    nzp_serv = 25;
                                    nzp_cnt = 1;
                                    break;
                                }
                            case "ЭЛН":
                                {
                                    nzp_serv = 210;
                                    nzp_cnt = 7;
                                    break;
                                }
                            default:
                                {
                                    nzp_serv = 0;
                                    nzp_cnt = 0;
                                    break;
                                }
                        }
                        if (nzp_serv == 0)
                        {
                            wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Yellow;
                            wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Yellow;
                            wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Yellow;
                            wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Yellow;
                            wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Yellow;
                            wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Yellow;
                        }
                        else
                        {
                            wb2.Worksheet(1).Row(i).Cell(3).Value = nzp_serv;
                        }
                    }
                }
                wb2.Save();
            }
            #endregion

            #region 110
            else if (type == 110)
            {
                var wb2 = new XLWorkbook(@"C:\Temp\Выписка50.xlsx");
                string address = "";
                Int32 nzp_serv;
                Int32 nzp_cnt;
                string[] separator = new string[] { "кв." };
                string[] separator1 = new string[] { "неж." };
                List<string> kvarParams = new List<string>();
                List<string> doubleKvars = new List<string>();
                Console.Write("Введите наименование БД:");
                string database = Console.ReadLine();
                for (int i = 1; i <= 18907; i++)
                {
                    if(i % 500 == 0)
                        Console.WriteLine(i.ToString());
                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(12).Value).Trim().Contains("Революционная"))
                    {
                        address = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(12).Value).Trim();
                        if (address.Contains("неж."))
                        {
                            kvarParams =
                                pg.SelectKvarParams("", database, address.Split(separator1, StringSplitOptions.None)[1].Trim(),
                                    7155107);
                        }
                        else
                        {
                            kvarParams = pg.SelectKvarParams("", database, 
                                address.Split(separator, StringSplitOptions.None)[1].Trim(), 7155107);
                        }
                    }
                    else
                    {
                        continue;
                    }
                    if (kvarParams == null)
                    {
                        wb2.Worksheet(1).Row(i).Cell(12).Style.Fill.BackgroundColor = XLColor.Red;
                    }
                    else
                    {
                        string month = "";
                        decimal saldo = 0;
                        bool august = false;
                        while (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim() != "ИТОГО:")
                        {
                            if (i % 500 == 0)
                                Console.WriteLine(i.ToString());
                            month = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim();
                            try
                            {
                                saldo = Math.Round(Convert.ToDecimal(wb2.Worksheet(1).Row(i).Cell(31).Value), 2);
                            }
                            catch
                            {
                                saldo = 0;
                            }
                            
                            if (month == "Август 2015")
                            {
                                if (saldo != 0)
                                {
                                    int t = 0;
                                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value) != "")
                                    {
                                        t = pg.InsertOutSaldo(kvarParams[0], kvarParams[1],
                                        100015, Convert.ToDecimal(wb2.Worksheet(1).Row(i).Cell(6).Value));
                                        if (t == 0)
                                        {
                                            wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Red;
                                        }
                                        else if (t == 2)
                                        {
                                            doubleKvars.Add(kvarParams[0]);
                                            august = true;
                                            break;
                                        }
                                    }
                                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(9).Value) != "")
                                    {
                                        t = pg.InsertOutSaldo(kvarParams[0], kvarParams[1],
                                        2, Convert.ToDecimal(wb2.Worksheet(1).Row(i).Cell(9).Value));
                                        if (t == 0)
                                        {
                                            wb2.Worksheet(1).Row(i).Cell(9).Style.Fill.BackgroundColor = XLColor.Red;
                                        }
                                    }
                                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(14).Value) != "")
                                    {
                                        t = pg.InsertOutSaldo(kvarParams[0], kvarParams[1],
                                        6, Convert.ToDecimal(wb2.Worksheet(1).Row(i).Cell(14).Value));
                                        if (t == 0)
                                        {
                                            wb2.Worksheet(1).Row(i).Cell(14).Style.Fill.BackgroundColor = XLColor.Red;
                                        }
                                    }
                                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(18).Value) != "")
                                    {
                                        t = pg.InsertOutSaldo(kvarParams[0], kvarParams[1],
                                        14, Convert.ToDecimal(wb2.Worksheet(1).Row(i).Cell(18).Value));
                                        if (t == 0)
                                        {
                                            wb2.Worksheet(1).Row(i).Cell(18).Style.Fill.BackgroundColor = XLColor.Red;
                                        }
                                    }
                                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(19).Value) != "")
                                    {
                                        t = pg.InsertOutSaldo(kvarParams[0], kvarParams[1],
                                        9, Convert.ToDecimal(wb2.Worksheet(1).Row(i).Cell(19).Value));
                                        if (t == 0)
                                        {
                                            wb2.Worksheet(1).Row(i).Cell(19).Style.Fill.BackgroundColor = XLColor.Red;
                                        }
                                    }
                                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(22).Value) != "")
                                    {
                                        t = pg.InsertOutSaldo(kvarParams[0], kvarParams[1],
                                        8, Convert.ToDecimal(wb2.Worksheet(1).Row(i).Cell(22).Value));
                                        if (t == 0)
                                        {
                                            wb2.Worksheet(1).Row(i).Cell(22).Style.Fill.BackgroundColor = XLColor.Red;
                                        }
                                    }
                                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(24).Value) != "")
                                    {
                                        t = pg.InsertOutSaldo(kvarParams[0], kvarParams[1],
                                        25, Convert.ToDecimal(wb2.Worksheet(1).Row(i).Cell(24).Value));
                                        if (t == 0)
                                        {
                                            wb2.Worksheet(1).Row(i).Cell(24).Style.Fill.BackgroundColor = XLColor.Red;
                                        }
                                    }
                                    
                                }
                                august = true;
                                break;
                            }
                            i++;
                        }
                        try
                        {
                            saldo = Math.Round(Convert.ToDecimal(wb2.Worksheet(1).Row(i - 1).Cell(31).Value), 2);
                        }
                        catch
                        {
                            saldo = 0;
                        }
                        if (saldo != 0 && !august)
                        {
                            int t = pg.InsertOutSaldo(kvarParams[0], kvarParams[1],
                                        100015, Convert.ToDecimal(wb2.Worksheet(1).Row(i - 1).Cell(6).Value));
                            if (t == 0)
                            {
                                wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Red;
                                t = pg.InsertOutSaldo(kvarParams[0], kvarParams[1],
                                2, Convert.ToDecimal(wb2.Worksheet(1).Row(i - 1).Cell(9).Value));
                                if (t == 0)
                                {
                                    wb2.Worksheet(1).Row(i).Cell(9).Style.Fill.BackgroundColor = XLColor.Red;
                                }
                                t = pg.InsertOutSaldo(kvarParams[0], kvarParams[1],
                                    6, Convert.ToDecimal(wb2.Worksheet(1).Row(i - 1).Cell(14).Value));
                                if (t == 0)
                                {
                                    wb2.Worksheet(1).Row(i).Cell(14).Style.Fill.BackgroundColor = XLColor.Red;
                                }
                                t = pg.InsertOutSaldo(kvarParams[0], kvarParams[1],
                                    14, Convert.ToDecimal(wb2.Worksheet(1).Row(i - 1).Cell(18).Value));
                                if (t == 0)
                                {
                                    wb2.Worksheet(1).Row(i).Cell(18).Style.Fill.BackgroundColor = XLColor.Red;
                                }
                                t = pg.InsertOutSaldo(kvarParams[0], kvarParams[1],
                                    9, Convert.ToDecimal(wb2.Worksheet(1).Row(i - 1).Cell(19).Value));
                                if (t == 0)
                                {
                                    wb2.Worksheet(1).Row(i).Cell(19).Style.Fill.BackgroundColor = XLColor.Red;
                                }
                                t = pg.InsertOutSaldo(kvarParams[0], kvarParams[1],
                                    8, Convert.ToDecimal(wb2.Worksheet(1).Row(i - 1).Cell(22).Value));
                                if (t == 0)
                                {
                                    wb2.Worksheet(1).Row(i).Cell(22).Style.Fill.BackgroundColor = XLColor.Red;
                                }
                                t = pg.InsertOutSaldo(kvarParams[0], kvarParams[1],
                                    25, Convert.ToDecimal(wb2.Worksheet(1).Row(i - 1).Cell(24).Value));
                                if (t == 0)
                                {
                                    wb2.Worksheet(1).Row(i).Cell(24).Style.Fill.BackgroundColor = XLColor.Red;
                                }
                            }
                            else if (t == 2)
                            {
                                doubleKvars.Add(kvarParams[0]);
                            }
                        }
                    }

                }
                StreamWriter errorRow = new StreamWriter(@"C:\Temp\outSaldo50.txt", false, Encoding.Default);
                foreach (string nzp_kvar in doubleKvars)
                {
                    errorRow.WriteLine(nzp_kvar);
                    pg.DelSaldo(nzp_kvar);
                }
                errorRow.Close();
                wb2.Save();
            }
            #endregion

            //Из Excel дата аоверки
            #region 111
            else if (type == 111)
            {
                var wb2 = new XLWorkbook(@"C:\Temp\14.09.2015(2).xlsx");
                string address = "";
                Int32 nzp_serv;
                Int32 nzp_cnt;
                int list_num = 0;
                int nzp_dom = 0;
                Console.Write("Введите наименование БД:");
                string database = Console.ReadLine();
                Console.Write("Введите номер листа:");
                list_num = Convert.ToInt32(Console.ReadLine());
                Console.Write("Введите id дома:");
                nzp_dom = Convert.ToInt32(Console.ReadLine());
                List<string> kvarParams = new List<string>();
                for (int i = 2; i <= 1000; i++)
                {
                    if (Convert.ToString(wb2.Worksheet(list_num).Row(i).Cell(1).Value).Trim() == "")
                        break;
                    if (Convert.ToString(wb2.Worksheet(list_num).Row(i).Cell(3).Value).Trim() != address
                        && Convert.ToString(wb2.Worksheet(list_num).Row(i).Cell(3).Value).Trim() != "")
                    {
                        address = Convert.ToString(wb2.Worksheet(list_num).Row(i).Cell(3).Value).Trim();
                        if (address.Contains("нежи."))
                        {
                            kvarParams = pg.SelectKvarParams("", database, address, nzp_dom);//7155105
                        }
                        else
                        {
                            kvarParams = pg.SelectKvarParams("", database, address, nzp_dom);
                        }

                    }
                    if (kvarParams == null)
                    {
                        wb2.Worksheet(list_num).Row(i).Cell(1).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(list_num).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(list_num).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(list_num).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(list_num).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(list_num).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(list_num).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                    else
                    {
                        String serv = Convert.ToString(wb2.Worksheet(list_num).Row(i).Cell(5).Value).Trim();
                        switch (serv)
                        {
                            case "Хим. очищенная вода":
                                {
                                    nzp_serv = 9;
                                    break;
                                }
                            case "Холодная вода":
                                {
                                    nzp_serv = 6;
                                    break;
                                }
                            default:
                                {
                                    nzp_serv = 0;
                                    break;
                                }
                        }
                        if (nzp_serv == 0)
                        {
                            wb2.Worksheet(list_num).Row(i).Cell(1).Style.Fill.BackgroundColor = XLColor.Blue;
                            wb2.Worksheet(list_num).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Blue;
                            wb2.Worksheet(list_num).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Blue;
                            wb2.Worksheet(list_num).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Blue;
                            wb2.Worksheet(list_num).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Blue;
                            wb2.Worksheet(list_num).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Blue;
                            wb2.Worksheet(list_num).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Blue;
                        }
                        else
                        {
                            int t = pg.UpdateCountersDatClose(database, kvarParams[0],
                                nzp_serv, 
                                Convert.ToString(wb2.Worksheet(list_num).Row(i).Cell(4).Value).Trim().Substring(1),
                                Convert.ToString(wb2.Worksheet(list_num).Row(i).Cell(9).Value).Trim());
                            if (t == 2)
                            {
                                wb2.Worksheet(list_num).Row(i).Cell(1).Style.Fill.BackgroundColor = XLColor.LightGreen;
                                wb2.Worksheet(list_num).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.LightGreen;
                                wb2.Worksheet(list_num).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.LightGreen;
                                wb2.Worksheet(list_num).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.LightGreen;
                                wb2.Worksheet(list_num).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.LightGreen;
                                wb2.Worksheet(list_num).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.LightGreen;
                                wb2.Worksheet(list_num).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.LightGreen;
                            }
                            else if (t == 0)
                            {
                                wb2.Worksheet(list_num).Row(i).Cell(1).Style.Fill.BackgroundColor = XLColor.Red;
                                wb2.Worksheet(list_num).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Red;
                                wb2.Worksheet(list_num).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Red;
                                wb2.Worksheet(list_num).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Red;
                                wb2.Worksheet(list_num).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Red;
                                wb2.Worksheet(list_num).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Red;
                                wb2.Worksheet(list_num).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Red;
                            }
                        }
                    }
                }
                wb2.Save();
            }
            #endregion

            //Из Excel формируем список параметров для записи в БД Билилнг
            #region 112
            else if (type == 112)
            {
                var wb2 = new XLWorkbook(@"C:\temp\февраль 2016.xlsx");
                string address = "";
                Int32 nzp_serv;
                Int32 nzp_cnt;
                Int32 start, end, nzp_dom;
                string[] separator = new string[] { "кв." };
                string[] separator1 = new string[] { "неж." };
                Console.Write("Введите наименование БД:");
                string database = Console.ReadLine();
                Console.Write("Введите ip:");
                string connStr = Console.ReadLine();
                Console.Write("Введите начальную строку:");
                start = Convert.ToInt32(Console.ReadLine());
                Console.Write("Введите конечную строку:");
                end = Convert.ToInt32(Console.ReadLine());
                Console.Write("Введите id дома:");
                nzp_dom = Convert.ToInt32(Console.ReadLine());
                List<string> kvarParams = new List<string>();
                for (int i = start; i <= end; i++)
                {
                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim() != address
                        && Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim() != "")
                    {
                        address = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim();
                        kvarParams = pg.SelectKvarParams(connStr, database, address, nzp_dom);//7155105
                    }
                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim() == "")
                    {
                        wb2.Worksheet(1).Row(i).Cell(2).Value = address;
                    }
                    if (kvarParams == null)
                    {
                        wb2.Worksheet(1).Row(i).Cell(1).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                    else
                    {
                        String serv = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim();
                        switch (serv)
                        {
                            case "ГВС":
                                {
                                    nzp_serv = 9;
                                    nzp_cnt = 3;
                                    break;
                                }
                            case "Водоснабжение":
                                {
                                    nzp_serv = 6;
                                    nzp_cnt = 2;
                                    break;
                                }
                            case "ЭЛ":
                            case "ЭЛД":
                            case "Э/энергия":
                                {
                                    nzp_serv = 25;
                                    nzp_cnt = 1;
                                    break;
                                }
                            case "ЭЛН":
                                {
                                    nzp_serv = 210;
                                    nzp_cnt = 7;
                                    break;
                                }
                            default:
                                {
                                    nzp_serv = 0;
                                    nzp_cnt = 0;
                                    break;
                                }
                        }
                        if (nzp_serv == 0)
                        {
                            wb2.Worksheet(1).Row(i).Cell(1).Style.Fill.BackgroundColor = XLColor.Yellow;
                            wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Yellow;
                            wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Yellow;
                            wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Yellow;
                            wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Yellow;
                        }
                        else
                        {
                            int nzp_counter = 0;
                            if (nzp_counter == 0)
                                nzp_counter = pg.InsertCounter(kvarParams[0], nzp_serv, Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim(), nzp_cnt, 
                                    Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value).Trim());
                            int t = pg.InsertCounterVal("", database, kvarParams[0], kvarParams[1], nzp_serv, Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim(),
                                "03.09.2015", Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value).Trim(), nzp_counter);
                            if (t == 0)
                            {
                                wb2.Worksheet(1).Row(i).Cell(1).Style.Fill.BackgroundColor = XLColor.Yellow;
                                wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Yellow;
                                wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Yellow;
                                wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Yellow;
                                wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Yellow;
                            }
                        }
                    }
                }
                wb2.Save();
            }
            #endregion

            //Из Excel формируем список параметров для записи в БД Билилнг
            #region 113
            else if (type == 113)
            {
                var wb2 = new XLWorkbook(@"C:\Temp\февраль 2016 — копия.xlsx");
                string address = "";
                Int32 nzp_serv;
                Int32 nzp_cnt;
                Int32 start, end, nzp_dom;
                string[] separator = new string[] { "кв." };
                string[] separator2 = new string[] { "комн." };
                string[] separator1 = new string[] { "неж." };
                Console.Write("Введите наименование БД:");
                string database = Console.ReadLine();
                Console.Write("Введите ip:");
                string connStr = Console.ReadLine();
                Console.Write("Введите начальную строку:");
                start = Convert.ToInt32(Console.ReadLine());
                Console.Write("Введите конечную строку:");
                end = Convert.ToInt32(Console.ReadLine());
                Console.Write("Введите id дома:");
                nzp_dom = Convert.ToInt32(Console.ReadLine());
                List<string> kvarParams = new List<string>();
                for (int i = start; i <= end; i++)
                {
                    try
                    {
                        if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim() != address
                        && Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim() != "")
                        {
                            //int nzp_dom = 0;
                            //if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim().Contains("40"))
                            //    nzp_dom = 7155105;
                            //else if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim().Contains("44"))
                            //    nzp_dom = 7155106;
                            //else if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim().Contains("50"))
                            //    nzp_dom = 7155107;
                            address = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim();
                            if (address.Contains("нежи."))
                            {
                                kvarParams = pg.SelectKvarParams(connStr, database,
                                    address.Split(separator1, StringSplitOptions.None)[1].Trim(), nzp_dom);//7155105
                            }
                            else
                            {
                                kvarParams = pg.SelectKvarParams(connStr, database, address.Split(separator, StringSplitOptions.None)[1]
                                    .Split(separator2, StringSplitOptions.None)[0].Trim(), nzp_dom);
                            }

                        }
                        if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim() == "")
                        {
                            wb2.Worksheet(1).Row(i).Cell(2).Value = address;
                        }
                        if (kvarParams == null)
                        {
                            wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Brown;
                            wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Brown;
                            wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Brown;
                            wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Brown;
                            wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Brown;
                            wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Brown;
                        }
                        else
                        {
                            String serv = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim();
                            switch (serv)
                            {
                                case "Горячая вода":
                                    {
                                        nzp_serv = 9;
                                        nzp_cnt = 3;
                                        break;
                                    }
                                case "Холодная вода":
                                    {
                                        nzp_serv = 6;
                                        nzp_cnt = 2;
                                        break;
                                    }
                                case "Электроснабжение":
                                case "ЭЛД":
                                    {
                                        nzp_serv = 25;
                                        nzp_cnt = 1;
                                        break;
                                    }
                                case "Электроснабжение ночное":
                                    {
                                        nzp_serv = 210;
                                        nzp_cnt = 7;
                                        break;
                                    }
                                default:
                                    {
                                        nzp_serv = 0;
                                        nzp_cnt = 0;
                                        break;
                                    }
                            }
                            if (nzp_serv == 0)
                            {
                                wb2.Worksheet(1).Row(i).Cell(1).Style.Fill.BackgroundColor = XLColor.Green;
                                wb2.Worksheet(1).Row(i).Cell(8).Style.Fill.BackgroundColor = XLColor.Green;
                                wb2.Worksheet(1).Row(i).Cell(9).Style.Fill.BackgroundColor = XLColor.Green;
                                wb2.Worksheet(1).Row(i).Cell(10).Style.Fill.BackgroundColor = XLColor.Green;
                            }
                            else
                            {
                                int nzp_counter = pg.SelectNzpCounter(connStr, database, kvarParams[0], nzp_serv, Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim()
                                    , "2016-03-01", Convert.ToString(wb2.Worksheet(1).Row(i).Cell(9).Value).Trim(), Convert.ToString(wb2.Worksheet(1).Row(i).Cell(13).Value).Trim());
                                if (nzp_counter != 0)
                                {
                                    //int t = pg.InsertCounterVal(connStr, database, kvarParams[0], kvarParams[1], nzp_serv,
                                    //    Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim(),
                                    //    "01.03.2016", Convert.ToString(wb2.Worksheet(1).Row(i).Cell(13).Value).Trim(),
                                    //    nzp_counter);
                                    //if (t == 0)
                                    //{
                                    //    wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Red;
                                    //    wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Red;
                                    //    wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Red;
                                    //    wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Red;
                                    //    wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Red;
                                    //    wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Red;
                                    //}
                                }
                                else
                                {
                                    wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Blue;
                                    wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Blue;
                                    wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Blue;
                                    wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Blue;
                                    wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Blue;
                                    wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Blue;
                                }

                                //t = pg.InsertCounterVal(kvarParams[0], kvarParams[1], nzp_serv, Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim(),
                                //    "01.08.2015", Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value).Trim(), nzp_counter);
                                //if (t == 0)
                                //{
                                //    wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Yellow;
                                //    wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Yellow;
                                //    wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Yellow;
                                //    wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Yellow;
                                //    wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Yellow;
                                //    wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Yellow;
                                //}
                                //t = pg.InsertCounterVal(kvarParams[0], kvarParams[1], nzp_serv, Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim(),
                                //    "01.06.2015", Convert.ToString(wb2.Worksheet(1).Row(i).Cell(7).Value).Trim(), nzp_counter);
                                //if (t == 0)
                                //{
                                //    wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Yellow;
                                //    wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Yellow;
                                //    wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Yellow;
                                //    wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Yellow;
                                //    wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Yellow;
                                //    wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Yellow;
                                //}
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        wb2.Worksheet(1).Row(i).Cell(12).Value = e.ToString();
                        wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.AshGrey;
                        wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.AshGrey;
                        wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.AshGrey;
                        wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.AshGrey;
                        wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.AshGrey;
                        wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.AshGrey;
                    }
                    
                }
                wb2.Save();
            }
            #endregion

            //Из Excel формируем список параметров для записи в БД Билилнг
            #region 1130
            else if (type == 1130)
            {
                var wb2 = new XLWorkbook(@"C:\Temp\Копия template_ipu0(1)-1.xlsx");
                string address = "";
                string kvar = "";
                Int32 nzp_serv;
                Int32 nzp_cnt;
                Int32 end;
                string[] separator = new string[] { "кв." };
                string[] separator1 = new string[] { "неж." };
                Console.Write("Введите наименование БД:");
                string database = Console.ReadLine();
                Console.Write("Введите ip:");
                string connStr = Console.ReadLine();
                Console.Write("Введите конечную строку:");
                end = Convert.ToInt32(Console.ReadLine());
                List<string> kvarParams = new List<string>();
                for (int i = 5; i <= end; i++)
                {
                    try
                    {
                        if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim() != "")
                        {
                            int nzp_dom = 0;
                            if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim().Contains("40"))
                                nzp_dom = 7155105;
                            else if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim().Contains("44"))
                                nzp_dom = 7155106;
                            else if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim().Contains("50"))
                                nzp_dom = 7155107;
                            address = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim();
                            if (address.Contains("нежи."))
                            {
                                kvarParams = pg.SelectKvarParams(connStr, database, Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim().Split(separator1, StringSplitOptions.None)[1].Trim(), nzp_dom);
                            }
                            else
                            {
                                kvarParams = pg.SelectKvarParams(connStr, database,
                                    Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value)
                                        .Trim()
                                        .Split(separator, StringSplitOptions.None)[1].Trim(), nzp_dom);
                            }

                        }
                        if (kvarParams == null)
                        {
                            wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Brown;
                            wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Brown;
                            wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Brown;
                            wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Brown;
                            wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Brown;
                            wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Brown;
                        }
                        else
                        {
                            String serv = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(7).Value).Trim();
                            switch (serv)
                            {
                                case "Горячая вода":
                                    {
                                        nzp_serv = 9;
                                        nzp_cnt = 3;
                                        break;
                                    }
                                case "Холодная вода":
                                    {
                                        nzp_serv = 6;
                                        nzp_cnt = 2;
                                        break;
                                    }
                                case "Электроснабжение":
                                case "ЭЛД":
                                    {
                                        nzp_serv = 25;
                                        nzp_cnt = 1;
                                        break;
                                    }
                                case "Электроснабжение ночное":
                                    {
                                        nzp_serv = 210;
                                        nzp_cnt = 7;
                                        break;
                                    }
                                default:
                                    {
                                        nzp_serv = 0;
                                        nzp_cnt = 0;
                                        break;
                                    }
                            }
                            if (nzp_serv == 0)
                            {
                                wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Green;
                                wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Green;
                                wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Green;
                                wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Green;
                                wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Green;
                                wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Green;
                            }
                            else
                            {
                                int nzp_counter = pg.SelectNzpCounter(connStr, database, kvarParams[0], nzp_serv,
                                    Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value).Trim(), "2016-03-01",
                                    Convert.ToString(wb2.Worksheet(1).Row(i).Cell(9).Value).Trim());
                                if (nzp_counter != 0)
                                {
                                    int t = pg.InsertCounterVal(connStr, database, kvarParams[0], kvarParams[1], nzp_serv, Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value).Trim(),
                                        "01.03.2016", Convert.ToString(wb2.Worksheet(1).Row(i).Cell(9).Value).Trim(), nzp_counter);
                                    if (t == 0)
                                    {
                                        wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Red;
                                        wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Red;
                                        wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Red;
                                        wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Red;
                                        wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Red;
                                        wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Red;
                                    }
                                    else
                                    {
                                        Console.WriteLine("строка =" + i + " загружена");
                                    }
                                        
                                    
                                }
                                else
                                {
                                    wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Blue;
                                    wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Blue;
                                    wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Blue;
                                    wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Blue;
                                    wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Blue;
                                    wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Blue;
                                }
                                //t = pg.InsertCounterVal(kvarParams[0], kvarParams[1], nzp_serv, Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim(),
                                //    "01.08.2015", Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value).Trim(), nzp_counter);
                                //if (t == 0)
                                //{
                                //    wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Yellow;
                                //    wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Yellow;
                                //    wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Yellow;
                                //    wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Yellow;
                                //    wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Yellow;
                                //    wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Yellow;
                                //}
                                //t = pg.InsertCounterVal(kvarParams[0], kvarParams[1], nzp_serv, Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim(),
                                //    "01.06.2015", Convert.ToString(wb2.Worksheet(1).Row(i).Cell(7).Value).Trim(), nzp_counter);
                                //if (t == 0)
                                //{
                                //    wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Yellow;
                                //    wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Yellow;
                                //    wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Yellow;
                                //    wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Yellow;
                                //    wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Yellow;
                                //    wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Yellow;
                                //}
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        wb2.Worksheet(1).Row(i).Cell(12).Value = e.ToString();
                        wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.AshGrey;
                        wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.AshGrey;
                        wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.AshGrey;
                        wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.AshGrey;
                        wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.AshGrey;
                        wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.AshGrey;
                    }
                    
                }
                wb2.Save();
            }
            #endregion

            #region 114
            else if (type == 114)
            {
                var wb2 = new XLWorkbook(@"C:\Temp\Начисление по услугам 40,44,50.xlsx");
                string address = "";
                Int32 nzp_serv;
                Int32 nzp_supp;
                string[] separator = new string[] { "кв." };
                string[] separator1 = new string[] { "неж." };
                List<string> kvarParams = new List<string>();
                List<string> doubleKvars = new List<string>();
                Console.Write("Введите наименование БД:");
                string database = Console.ReadLine();
                for (int i = 4; i <= 163; i++)
                {
                    if(i%500 == 0)
                        Console.WriteLine(i);
                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim() != address
                        && Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim() != "")
                    {
                        address = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim();
                        if (address.Contains("нежи."))
                        {
                            kvarParams =
                                pg.SelectKvarParams("", database, address.Split(separator1, StringSplitOptions.None)[1].Trim(),
                                    7155106);
                        }
                        else
                        {
                            int nzp_dom = 0;
                            if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim().Contains("д.40"))
                                nzp_dom = 7155105;
                            else if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim().Contains("д.44"))
                                nzp_dom = 7155106;
                            else if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim().Contains("д.50"))
                                nzp_dom = 7155107;
                            kvarParams = pg.SelectKvarParamsByNumLs(
                                Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim().Substring(5), nzp_dom);
                        }

                    }
                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim() == "")
                    {
                        wb2.Worksheet(1).Row(i).Cell(2).Value = address;
                    }
                    if (kvarParams == null)
                    {
                        wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                    else
                    {
                        String serv = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value).Trim();
                        switch (serv)
                        {
                            case "Перерасход Водоотведение":
                            case "Водоотведение":
                                {
                                    nzp_serv = 7;
                                    nzp_supp = 101191;
                                    break;
                                }
                            case "Коррек. отопления":
                            case "Отопление":
                                {
                                    nzp_serv = 8;
                                    nzp_supp = 101191;
                                    break;
                                }
                            case "Электроэнергия":
                            case "Электроэнергия (день)":
                                {
                                    nzp_serv = 25;
                                    nzp_supp = 101192;
                                    break;
                                }
                            case "Электроэнергия МОП (день)":
                            case "Электроэнергия МОП":
                            {
                                nzp_serv = 515;
                                nzp_supp = 101192;
                                break;
                            }

                            case "Подогрев":
                                {
                                    nzp_serv = 14;
                                    nzp_supp = 101191;
                                    break;
                                }
                            case "Содержание":
                                {
                                    nzp_serv = 17;
                                    nzp_supp = 101190;
                                    break;
                                }
                            case "Капитальный ремонт":
                            case "Текущий ремонт":
                                {
                                    nzp_serv = 2;
                                    nzp_supp = 101190;
                                    break;
                                }
                            case "Хим. очищенная вода":
                                {
                                    nzp_serv = 9;
                                    nzp_supp = 101191;
                                    break;
                                }
                            case "Холодная вода":
                                {
                                    nzp_serv = 6;
                                    nzp_supp = 101191;
                                    break;
                                }
                            case "Электроэнергия (ночь)":
                                {
                                    nzp_serv = 210;
                                    nzp_supp = 101192;
                                    break;
                                }
                            case "Электроэнергия МОП (ночь)":
                            {
                                nzp_serv = 516;
                                nzp_supp = 101192;
                                break;
                            }

                            case "Домофон":
                                {
                                    nzp_serv = 26;
                                    nzp_supp = 101194;
                                    break;
                                }
                            case "Перерасход ГВС":
                                {
                                    nzp_serv = 514;
                                    nzp_supp = 101191;
                                    break;
                                }
                            case "Перерасход Подогрев":
                                {
                                    nzp_serv = 513;
                                    nzp_supp = 101191;
                                    break;
                                }
                            case "Перерасход ХВС":
                                {
                                    nzp_serv = 510;
                                    nzp_supp = 101191;
                                    break;
                                }
                            case "Разовое снятие":
                                {
                                    nzp_serv = 100021;
                                    nzp_supp = 101190;
                                    break;
                                }
                            case "Разовая услуга":
                                {
                                    nzp_serv = 100022;
                                    nzp_supp = 101190;
                                    break;
                                }
                            default:
                                {
                                    nzp_serv = 0;
                                    nzp_supp = 0;
                                    break;
                                }
                        }
                        if (nzp_serv == 0)
                        {
                            wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.LightGreen;
                            wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.LightGreen;
                            wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.LightGreen;
                            wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.LightGreen;
                            wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.LightGreen;
                            wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.LightGreen;
                        }
                        else
                        {
                            if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(8).Value).Trim() != ""
                                && Convert.ToDecimal(wb2.Worksheet(1).Row(i).Cell(8).Value) != 0)
                            {
                                var t = pg.InsertOutSaldo("billTlt", kvarParams[0], kvarParams[1],
                                    nzp_serv, Convert.ToDecimal(wb2.Worksheet(1).Row(i).Cell(8).Value), nzp_supp);
                                if (t == 0)
                                {
                                    wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Red;
                                    wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Red;
                                    wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Red;
                                    wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Red;
                                    wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Red;
                                    wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Red;
                                }
                                else if (t == 2)
                                {
                                    //doubleKvars.Add(kvarParams[0]);
                                    wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Orange;
                                    wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Orange;
                                    wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Orange;
                                    wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Orange;
                                    wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Orange;
                                    wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Orange;
                                }
                            }
                            else
                            {
                                wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.LightBlue;
                                wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.LightBlue;
                                wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.LightBlue;
                                wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.LightBlue;
                                wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.LightBlue;
                                wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.LightBlue;
                            }
                        }
                    }
                }
                foreach (string nzp_kvar in doubleKvars)
                {
                    pg.DelSaldo(nzp_kvar);
                }
                wb2.Save();
            }
            #endregion

            //Сальдо по пени
            #region 115
            else if (type == 115)
            {
                var wb2 = new XLWorkbook(@"C:\Temp\Копия Пени 50.xlsx");
                string address = "";
                Int32 nzp_serv;
                Int32 nzp_supp;
                List<string> kvarParams = new List<string>();
                List<string> doubleKvars = new List<string>();
                for (int i = 2; i <= 142; i++)
                {
                    if (doubleKvars.Contains(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim().Substring(5)))
                    {
                        wb2.Worksheet(1).Row(i).Cell(1).Style.Fill.BackgroundColor = XLColor.Red;
                        wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Red;
                        continue;
                    }
                    else
                    {
                        doubleKvars.Add(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim().Substring(5));
                    }

                    kvarParams =
                            pg.SelectNzpKvarByNumLs("billTlt", Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim().Substring(5));

                    if (kvarParams == null)
                    {
                        wb2.Worksheet(1).Row(i).Cell(1).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                    else
                    {
                        nzp_serv = 500;
                        nzp_supp = 101190;
                                
                        int t = pg.InsertOutSaldo("billTlt",
                                                    kvarParams[0], 
                                                        kvarParams[1],
                                                            nzp_serv, 
                                                                Convert.ToDecimal(wb2.Worksheet(1).Row(i).Cell(2).Value), 
                                                                    nzp_supp);
                        if (t == 0)
                        {
                            wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Red;
                            wb2.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Red;
                            wb2.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Red;
                            wb2.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Red;
                            wb2.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Red;
                            wb2.Worksheet(1).Row(i).Cell(7).Style.Fill.BackgroundColor = XLColor.Red;
                        }
                    }
                }
                wb2.Save();
            }
            #endregion

            //Сальдо по пени
            #region 116
            else if (type == 116)
            {
                var book = new XLWorkbook(@"C:\Temp\Реестр паспортиста по 7Просека 94.xlsx");
                string address = "";
                Int32 nzp_serv;
                Int32 nzp_supp;
                String nkvar = "";
                String nzp_kvar = "";
                List<string> kvarParams = new List<string>();
                List<string> doubleKvars = new List<string>();
                Boolean svid = false;
                for (int i = 35; i <= 84; i++)
                {
                    if (nkvar != Convert.ToString(book.Worksheet(1).Row(i).Cell(2).Value).Trim() && Convert.ToString(book.Worksheet(1).Row(i).Cell(2).Value).Trim() != "")
                    {
                        nzp_kvar = pg.SelectNzpKvarByKvarDom("billAuk",
                            Convert.ToString(book.Worksheet(1).Row(i).Cell(2).Value).Trim(), 7155105);
                        svid = true;
                    }
                    else
                    {
                        svid = false;
                    }
                    if (nzp_kvar.Split('|')[0] == "0")
                    {
                        book.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                        book.Worksheet(1).Row(i).Cell(21).Value = nzp_kvar.Split('|')[1];
                    }
                    else
                    {
                        int nzp_gil = pg.InsertGil();
                        int nzp_rod = 0;
                        #region nzp_rod
                        switch (Convert.ToString(book.Worksheet(1).Row(i).Cell(4).Value).Trim())
                        {
                            case "брат":
                                {
                                    nzp_rod = 587;
                                    break;
                                }
                            case "внук":
                                {
                                    nzp_rod = 560;
                                    break;
                                }
                            case "внучка":
                                {
                                    nzp_rod = 568;
                                    break;
                                }
                            case "гр.муж":
                                {
                                    nzp_rod = 571;
                                    break;
                                }
                            case "двоюродн.":
                                {
                                    nzp_rod = 15;
                                    break;
                                }
                            case "дочь":
                                {
                                    nzp_rod = 559;
                                    break;
                                }
                            case "дядя":
                                {
                                    nzp_rod = 738;
                                    break;
                                }
                            case "жена":
                                {
                                    nzp_rod = 562;
                                    break;
                                }
                            case "зять":
                                {
                                    nzp_rod = 565;
                                    break;
                                }
                            case "кс":
                                {
                                    nzp_rod = 561;
                                    break;
                                }
                            case "мать":
                                {
                                    nzp_rod = 563;
                                    break;
                                }
                            case "мать мужа":
                                {
                                    nzp_rod = 619;
                                    break;
                                }
                            case "муж":
                                {
                                    nzp_rod = 567;
                                    break;
                                }
                            case "отец":
                                {
                                    nzp_rod = 572;
                                    break;
                                }
                            case "отчим":
                                {
                                    nzp_rod = 640;
                                    break;
                                }
                            case "падчерица":
                                {
                                    nzp_rod = 30;
                                    break;
                                }
                            case "племянник":
                                {
                                    nzp_rod = 666;
                                    break;
                                }
                            case "племянница":
                                {
                                    nzp_rod = 899;
                                    break;
                                }
                            case "сестра":
                                {
                                    nzp_rod = 899;
                                    break;
                                }
                            case "сноха":
                                {
                                    nzp_rod = 575;
                                    break;
                                }
                            case "сын":
                                {
                                    nzp_rod = 564;
                                    break;
                                }
                            case "сын жены":
                                {
                                    nzp_rod = 1156;
                                    break;
                                }
                            case "тетя":
                                {
                                    nzp_rod = 596;
                                    break;
                                }
                            case "теща":
                                {
                                    nzp_rod = 594;
                                    break;
                                }
                            case "собств":
                            case "собств.":
                                {
                                    nzp_rod = 582;
                                    break;
                                }
                        }
                        #endregion
                        //int nzp_dok = 0;
                        //#region nzp_dok
                        //switch (Convert.ToString(book.Worksheet(1).Row(i).Cell(10).Value).Trim())
                        //{
                        //    case "паспорт":
                        //        {
                        //            nzp_dok = 10;
                        //            break;
                        //        }
                        //    case "Св-во о рожд.":
                        //        {
                        //            nzp_dok = 2;
                        //            break;
                        //        }
                        //    case "Св-во о рождении":
                        //        {
                        //            nzp_dok = 2;
                        //            break;
                        //        }
                        //    case "Св-во рожд.":
                        //        {
                        //            nzp_dok = 2;
                        //            break;
                        //        }
                        //    default:
                        //        {
                        //            nzp_dok = -1;
                        //            break;
                        //        }
                        //}
                        //#endregion

                        //string serij = "";
                        //if (Convert.ToString(book.Worksheet(1).Row(i).Cell(11).Value).Trim() != "" && Convert.ToString(book.Worksheet(1).Row(i).Cell(11).Value).Trim().Length >= 4)
                        //{
                        //    if (nzp_dok == 10)
                        //        serij = Convert.ToString(book.Worksheet(1).Row(i).Cell(11).Value).Trim().Substring(0, 2) + " " + Convert.ToString(book.Worksheet(1).Row(i).Cell(11).Value).Trim().Substring(2, 2);
                        //    else
                        //        serij = Convert.ToString(book.Worksheet(1).Row(i).Cell(11).Value).Trim();
                        //}
                        if (Convert.ToString(book.Worksheet(1).Row(i).Cell(5).Value).Trim().ToUpper() != "")
                        {
                            int nzp_kart = pg.InsertKart("billAuk", nzp_gil, nzp_kvar.Split('|')[0],
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(5).Value).Trim().ToUpper().Split(' ')[0],
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(5).Value).Trim().ToUpper().Split(' ')[1],
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(5).Value).Trim().ToUpper().Split(' ')[2],
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(6).Value).Trim(),
                                                    "",
                                                    (svid) ? Convert.ToString(book.Worksheet(1).Row(i + 1).Cell(10).Value).Trim() : "",
                                                    nzp_rod,
                                                    Convert.ToString(book.Worksheet(1).Row(i).Cell(4).Value).Trim());
                            pg.InsertGrgd(nzp_kart);
                        }
                        
                    }
                }
                book.Save();
            }
            #endregion

            #region ПАЧКИ ОПЛАТ ЕИРЦ
            #region 1000 Формирование оплат для NCC
            else if (type == 1000)
            {
                Int32 year = 2016;
                Int32 month = 2;
                var wb2 = new XLWorkbook(@"C:\temp\ЕИРЦ_Пачки\NCC_Out_num.xlsx");
                int num = Convert.ToInt32(wb2.Worksheet(1).Row(1).Cell(1).Value) + 1;
                wb2.Worksheet(1).Row(1).Cell(1).Value = num;
                wb2.Save();
                StreamWriter outPack = new StreamWriter(@"C:\Temp\ЕИРЦ_Пачки\ncc_out.txt", false, Encoding.GetEncoding("cp866"));
                outPack.WriteLine("<smpay_load_hdr><format_id>smpay_load_data</format_id><format_version>1</format_version><file_id>" + num + "</file_id></smpay_load_hdr>");
                List<string> prefs = new List<string>() {"bill01", "bill02"};
                DataTable dt = pg.SelectSaldoForNCC(year, month, prefs);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    const string uniqNumOrg = "00103198";
                    string id = uniqNumOrg + Convert.ToString(dt.Rows[i][0]).PadRight(20, ' ');
                    string address = Convert.ToString(dt.Rows[i][1]).PadRight(40, ' ');
                    decimal d = Math.Round(Convert.ToDecimal(dt.Rows[i][2]), 2);
                    int saldo = Convert.ToInt32(d * 100) > 0 ? Convert.ToInt32(d * 100) : 0;
                    outPack.WriteLine(id + address + saldo.ToString().PadLeft(12, '0'));
                    //outPack.WriteLine(id + address);
                }
                outPack.Close();
            }
            #endregion

            #region 1100 Формирование оплат для NCC
            else if (type == 1100)
            {
                Int32 year = 2016;
                Int32 month = 2;
                StreamWriter outPack = new StreamWriter(@"C:\Temp\dymok.csv", false, Encoding.UTF8);
                List<string> prefs = new List<string>() { "bill01", "bill02" };
                DataTable dt = pg.SelectSaldoForAvtovazbank(year, month, prefs);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string eirc = Convert.ToString(dt.Rows[i][4]);
                    string id = Convert.ToString(dt.Rows[i][0]);
                    string address = Convert.ToString(dt.Rows[i][1]);
                    string d = Math.Round(Convert.ToDecimal(dt.Rows[i][2]) > 0 ? Convert.ToDecimal(dt.Rows[i][2]) : 0, 2).ToString().Replace(',','.');
                    string fio = Convert.ToString(dt.Rows[i][3]);
                    string[] objsObj = fio.Split(new[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                    fio = "";
                    foreach (string word in objsObj.ToList<string>())
                    {
                        if (fio != "")
                        {
                            fio = fio + word.Substring(0, 1) + ". ";
                        }
                        else
                        {
                            fio = word.Substring(0, 1) + ". ";
                        }
                    }
                    outPack.WriteLine(eirc + id + ";" + address + ";" + d + ";" + fio);
                }
                outPack.Close();
            }
            #endregion

            #region 1200 Формирование оплат для Сбербанка
            else if (type == 1200)
            {
                String fileName = "6321388192_40702810754400005587_001_y01";
                StreamWriter outPack = new StreamWriter(@"C:\Temp\" + fileName + ".txt", false, Encoding.GetEncoding("windows-1251"));
                string db = "192.168.1.25";
                List<string> prefs = new List<string>() { "bill01", "bill02" };
                DataTable dt = pg.SelectSaldoForSberbank(db, 2, 2016, prefs);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string num_ls = Convert.ToString(dt.Rows[i][0]).PadLeft(6, '0');
                    string fio = Convert.ToString(dt.Rows[i][1]);
                    string address = Convert.ToString(dt.Rows[i][2]);
                    string period = Convert.ToString(dt.Rows[i][3]);
                    string d = Math.Round(Convert.ToDecimal(dt.Rows[i][4]), 2) >= 0 ? Math.Round(Convert.ToDecimal(dt.Rows[i][4]), 2).ToString().Replace(',', '.') : "0";

                    outPack.WriteLine(num_ls + ";" + fio + ";" + address + ";" + period + ";" + d);
                }
                outPack.Close();
            }
            #endregion

            #region 1300 Формирование оплат для NCC
            else if (type == 1300)
            {
                Int32 year = 2016;
                Int32 month = 2;
                StreamWriter outPack = new StreamWriter(@"C:\Temp\avtovazbank.csv", false, Encoding.UTF8);
                List<string> prefs = new List<string>() { "bill01", "bill02" };
                DataTable dt = pg.SelectSaldoForAvtovazbank2(year, month, prefs);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string id = Convert.ToString(dt.Rows[i][0]);
                    string address = Convert.ToString(dt.Rows[i][1]);
                    string fio = Convert.ToString(dt.Rows[i][3]);
                    string d = Math.Round(Convert.ToDecimal(dt.Rows[i][2]) > 0 ? Convert.ToDecimal(dt.Rows[i][2]) : 0, 2).ToString().Replace(',', '.');
                    
                    outPack.WriteLine(id + ";" + address + ";" + fio + ";" + d);
                }
                outPack.Close();
            }
            #endregion
            #endregion

            #region 1110
            else if (type == 1110)
            {
                Start:
                string name = Console.ReadLine();
                if (name == "ex")
                    return;
                StreamReader sr = new StreamReader(@"C:\Temp\AVBpack\" + name + ".txt", System.Text.Encoding.Default);
                string line;
                string[] separator = new string[] { ";" };
                int k = 0;
                StreamWriter пачка = new StreamWriter(@"C:\Temp\AVBpack\Return" + name + ".txt", false, Encoding.Default);
                List<string> lsStr = new List<string>();
                int count = 0;
                decimal sum = 0;
                int num = 1;
                string paidDate = "14.12.2015";
                string operDay = "15.12.2015";
                string packDate = "";
                while ((line = sr.ReadLine()) != null)
                {
                    string str = "";
                    string[] incomingData = line.Split(separator, StringSplitOptions.None);
                    var paid = incomingData[2].Trim().Replace('.',',');
                    decimal p = Convert.ToDecimal(paid);
                    string pkod = incomingData[1];
                    if (pkod != "00" && pkod != "0")
                    {
                        count++;
                        sum += p;
                        str += "@@@|" + num.ToString() + "||33|2|" + pkod + "|" +
                            Convert.ToDateTime(incomingData[0]).ToShortDateString() + "|" + paidDate + "|0000|0.00|"
                            + p.ToString().Replace(',', '.') + "|0|0|0|0.00|||";
                        lsStr.Add(str);
                        num++;
                        packDate = incomingData[0];
                    }
                }
                string str1 = "";
                string str2 = "";
                Random rnd = new Random();
                int packNum = rnd.Next(1, 1000000001); // creates a number between 1 and 12
                str1 += "***|АВБ||" + packNum.ToString() + "|" + Convert.ToDateTime(packDate).ToShortDateString() + "|12:00:00" +
                    "|" + operDay + "|1|0.00|" + sum.ToString().Replace(',', '.') + "|0|0.00|!1.00|";
                str2 += "###|АВБ||" + packNum.ToString() + "|" + Convert.ToDateTime(packDate).ToShortDateString() +
                    "|" + operDay + "|" + count.ToString() + "|0.00|" + sum.ToString().Replace(',', '.') + "|0|0.00|0|!1.00|";
                пачка.WriteLine(str1);
                пачка.WriteLine(str2);
                int listCount = 0;
                foreach (string str in lsStr)
                {
                    listCount++;
                    if (listCount == lsStr.Count)
                        пачка.Write(str);
                    else
                        пачка.WriteLine(str);
                }
                пачка.Close();
                goto Start;
            }
            #endregion

            #region 1111
            else if (type == 1111)
            {
            Start:
                string name = Console.ReadLine();
                if (name == "ex")
                    return;
                StreamReader sr = new StreamReader(@"C:\Temp\AVBpack\" + name + ".txt", System.Text.Encoding.Default);
                string line;
                string[] separator = new string[] { "\t" };
                int k = 0;
                StreamWriter пачка = new StreamWriter(@"C:\Temp\AVBpack\Return" + name + ".txt", false, Encoding.Default);
                List<string> lsStr = new List<string>();
                int count = 0;
                decimal sum = 0;
                int num = 1;
                string paidDate = "08.12.2015";
                string operDay = "08.12.2015";
                string packDate = "";
                while ((line = sr.ReadLine()) != null)
                {
                    string str = "";
                    string[] incomingData = line.Split(separator, StringSplitOptions.None);
                    var paid = incomingData[1].Trim().Replace('.', ',');
                    decimal p = Convert.ToDecimal(paid);
                    string pkod = pg.SelectPkodByNumLs("billTlt", incomingData[0].Substring(5))[0];
                    if (pkod != "00" && pkod != "0")
                    {
                        count++;
                        sum += p;
                        str += "@@@|" + num.ToString() + "||33|2|" + pkod + "|" +
                            Convert.ToDateTime(incomingData[2]).ToShortDateString() + "|" + paidDate + "|0000|0.00|"
                            + p.ToString().Replace(',', '.') + "|0|0|0|0.00|||";
                        lsStr.Add(str);
                        num++;
                        packDate = incomingData[2];
                    }
                }
                string str1 = "";
                string str2 = "";
                Random rnd = new Random();
                int packNum = rnd.Next(1, 1000000001); // creates a number between 1 and 12
                str1 += "***|АВБ||" + packNum.ToString() + "|" + Convert.ToDateTime(packDate).ToShortDateString() + "|12:00:00" +
                    "|" + operDay + "|1|0.00|" + sum.ToString().Replace(',', '.') + "|0|0.00|!1.00|";
                str2 += "###|АВБ||" + packNum.ToString() + "|" + Convert.ToDateTime(packDate).ToShortDateString() +
                    "|" + operDay + "|" + count.ToString() + "|0.00|" + sum.ToString().Replace(',', '.') + "|0|0.00|0|!1.00|";
                пачка.WriteLine(str1);
                пачка.WriteLine(str2);
                int listCount = 0;
                foreach (string str in lsStr)
                {
                    listCount++;
                    if (listCount == lsStr.Count)
                        пачка.Write(str);
                    else
                        пачка.WriteLine(str);
                }
                пачка.Close();
                goto Start;
            }
            #endregion

            #region 117
            else if (type == 117)
            {
                DataTable dt = ora.SelectKoap();
                var wb2 = new XLWorkbook(@"C:\Temp\2222.xlsx");

                for (int i = 8; i <= 1000; i++)//18953
                {
                    if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim() == "Итого по району")
                        continue;
                    if(i % 500 == 0)
                        Console.WriteLine(i);
                    try
                    {
                        for (int j = 0; j < dt.Rows.Count; j++)
                        {
                            if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim() ==
                                Convert.ToString(dt.Rows[j][0]))
                            {
                                switch (Convert.ToString(dt.Rows[j][1]))
                                {
                                    case "ч.1. ст 14.1.3 КоАП РФ":
                                    case "ч.1 ст. 14.1.3 КоАП РФ":
                                    {
                                        wb2.Worksheet(1).Row(i).Cell(5).Value = "1"; 
                                        break;
                                    }
                                    case "ч.2 ст. 14.1.3 КоАП РФ":
                                    {
                                        wb2.Worksheet(1).Row(i).Cell(6).Value = "1"; 
                                        break;
                                    }
                                    case "ч.24 ст 19.5 КоАП РФ":
                                    {
                                        wb2.Worksheet(1).Row(i).Cell(7).Value = "1"; 
                                        break;
                                    }
                                    case "ч.1 ст. 7.23.3 КоАП РФ":
                                    {
                                        wb2.Worksheet(1).Row(i).Cell(3).Value = "1";
                                        break;
                                    }
                                    case "ч.2 ст. 7.23.3 КоАП РФ":
                                    {
                                        wb2.Worksheet(1).Row(i).Cell(4).Value = "1";
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {

                    }

                }
                wb2.Save();
            }
            #endregion

            #region 118
            else if (type == 118)
            {

                var wb1 = new XLWorkbook(@"C:\Temp\8915.xlsx");
                var wb2 = new XLWorkbook(@"C:\Temp\generator.xlsx");

                for (int i = 1; i <= 269; i++)//18953
                {
                    try
                    {
                        for (int j = 1; j <= 269; j++)
                        {
                            if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim() ==
                                Convert.ToString(wb1.Worksheet(1).Row(j).Cell(1).Value).Trim())
                            {
                                string s1 = wb2.Worksheet(1).Row(i).Cell(2).Value.ToString().Replace('.', ',');
                                string s2 = wb1.Worksheet(1).Row(j).Cell(2).Value.ToString().Replace('.', ',');

                                decimal d1 = Convert.ToDecimal(s1);
                                decimal d2 = Convert.ToDecimal(s2);

                                string str = "33333";
                                if (d1 != d2)
                                {
                                    wb2.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Yellow;
                                    wb2.Worksheet(1).Row(i).Cell(1).Style.Fill.BackgroundColor = XLColor.Yellow;
                                }
                                else
                                {
                                    wb1.Worksheet(1).Row(j).Cell(2).Style.Fill.BackgroundColor = XLColor.Yellow;
                                    wb1.Worksheet(1).Row(j).Cell(1).Style.Fill.BackgroundColor = XLColor.Yellow;
                                    
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {

                    }

                }
                wb2.Save();
                wb1.Save();
            }
            #endregion

            #region 119
            else if (type == 119)
            {
                var wb1 = new XLWorkbook(@"C:\Temp\IMPORTTOEZHKH.xlsx");
                int gkhCode = 0;
                string address = "";
                for (int i = 1; i <= 3889; i++)//18953
                {
                    if (address == Convert.ToString(wb1.Worksheet(1).Row(i).Cell(1).Value).Trim())
                    {
                        if (gkhCode != 0)
                            wb1.Worksheet(1).Row(i).Cell(10).Value = gkhCode;
                    }
                    else
                    {
                        address = Convert.ToString(wb1.Worksheet(1).Row(i).Cell(1).Value).Trim();
                        gkhCode = ora.SelectGkhCode(address);
                        if (gkhCode != 0)
                            wb1.Worksheet(1).Row(i).Cell(10).Value = gkhCode;
                    }
                }
                wb1.Save();
            }
            #endregion

            #region 120 Сумма по двум эксель файлам
            else if (type == 120)
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
            #endregion

            #region 121 Делим значения на 2 в Эксель
            else if (type == 121)
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
            #endregion

            #region 122 Формируем скрипт из txt
            else if (type == 122)
            {
                StreamWriter script = new StreamWriter(@"C:\Temp\script.txt", false, Encoding.Default);
                for (int i = 1; i <= 12; i++)
                {
                    String line = "ALTER TABLE bill01_charge_16.fn_supplier" + i.ToString("00") +
                                  " DROP CONSTRAINT fk_fn_supplier" + i.ToString("00") + "_nzp_pack_ls;";
                    script.WriteLine(line);
                    line = "ALTER TABLE bill01_charge_16.fn_supplier" + i.ToString("00") +
                           " ADD CONSTRAINT fk_fn_supplier" + i.ToString("00") +
                           "_nzp_pack_ls FOREIGN KEY (nzp_pack_ls) REFERENCES fbill_fin_16.pack_ls (nzp_pack_ls) MATCH SIMPLE ON UPDATE NO ACTION ON DELETE NO ACTION;";
                    script.WriteLine(line);
                }
                for (int i = 1; i <= 12; i++)
                {
                    int dayInMonth = DateTime.DaysInMonth(2016, i);
                    for (int j = 1; j <= dayInMonth; j++)
                    {
                        String line = "ALTER TABLE fbill_fin_16.fn_pa_dom_" + i.ToString("00") +
                                  "_" + j.ToString("00") +
                                  " DROP CONSTRAINT cnstr_fn_pa_dom_" + i.ToString("00") +
                                  "_" + j.ToString("00") +
                                  "_dat_oper;";
                        script.WriteLine(line);
                        line = "ALTER TABLE fbill_fin_16.fn_pa_dom_" + i.ToString("00") +
                                  "_" + j.ToString("00") +
                                  " ADD CONSTRAINT cnstr_fn_pa_dom_" + i.ToString("00") +
                                  "_" + j.ToString("00") +
                                  "_dat_oper CHECK (dat_oper = '2016-" + i.ToString("00") +
                                  "-" + j.ToString("00") +
                                  "'::date);";
                        script.WriteLine(line);
                    }
                }
                for (int i = 1; i <= 11; i++)
                {
                    String line = "ALTER TABLE fbill_fin_16.fn_perc_dom_" + i.ToString("00") +
                                  " DROP CONSTRAINT cnstr_fn_perc_dom_" + i.ToString("00") +
                                  "_dat_oper;";
                    script.WriteLine(line);
                    line = "ALTER TABLE fbill_fin_16.fn_perc_dom_" + i.ToString("00") +
                                  " ADD CONSTRAINT cnstr_fn_perc_dom_" + i.ToString("00") +
                                  "_dat_oper CHECK (dat_oper >= '2016-" + i.ToString("00") +
                                  "-01'::date AND dat_oper < '2016-" + (i+1).ToString("00") +
                                  "-01'::date);";
                    script.WriteLine(line);
                }
                for (int i = 1; i <= 11; i++)
                {
                    String line = "ALTER TABLE fbill_fin_16.fn_naud_dom_" + i.ToString("00") +
                                  " DROP CONSTRAINT cnstr_fn_naud_dom_" + i.ToString("00") +
                                  "_dat_oper;";
                    script.WriteLine(line);
                    line = "ALTER TABLE fbill_fin_16.fn_naud_dom_" + i.ToString("00") +
                                  " ADD CONSTRAINT cnstr_fn_naud_dom_" + i.ToString("00") +
                                  "_dat_oper CHECK (dat_oper >= '2016-" + i.ToString("00") +
                                  "-01'::date AND dat_oper < '2016-" + (i + 1).ToString("00") +
                                  "-01'::date);";
                    script.WriteLine(line);
                }
                for (int i = 1; i <= 12; i++)
                {
                    int dayInMonth = DateTime.DaysInMonth(2016, i);
                    for (int j = 1; j <= dayInMonth; j++)
                    {
                        String line = "ALTER TABLE fbill_fin_16.fn_distrib_dom_" + i.ToString("00") +
                                  "_" + j.ToString("00") +
                                  " DROP CONSTRAINT cnstr_fn_distrib_dom_" + i.ToString("00") +
                                  "_" + j.ToString("00") +
                                  "_dat_oper;";
                        script.WriteLine(line);
                        line = "ALTER TABLE fbill_fin_16.fn_distrib_dom_" + i.ToString("00") +
                                  "_" + j.ToString("00") +
                                  " ADD CONSTRAINT cnstr_fn_distrib_dom_" + i.ToString("00") +
                                  "_" + j.ToString("00") +
                                  "_dat_oper CHECK (dat_oper = '2016-" + i.ToString("00") +
                                  "-" + j.ToString("00") +
                                  "'::date);";
                        script.WriteLine(line);
                    }
                }
                script.Close();
            }
            #endregion

            #region 123 Пробегаем по всем таблицам в БД
            else if (type == 123)
            {
                var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("1");
                var tables = ora.SelectTableFromDB();
                int row = 2;
                for (int i = 0; i < tables.Rows.Count; i++)
                {
                    var rowsCount = ora.SelectRowsCount("base", tables.Rows[i][0].ToString());
                    var rowsCountTest = ora.SelectRowsCount("test", tables.Rows[i][0].ToString());
                    ws.Cell(row, 1).Value = tables.Rows[i][0].ToString();
                    ws.Cell(row, 2).Value = rowsCount;
                    ws.Cell(row, 3).Value = rowsCountTest;
                    ws.Cell(row, 4).Value = rowsCount - rowsCountTest;
                    row++;
                }
               
                wb.SaveAs(@"C:\temp\EzhkhDBRowsCount.xlsx");
            }
            #endregion

            #region 124 MaxId from db
            else if (type == 124)
            {
                var tables = ora.SelectTableFromDB();
                Int32 maxId = 0;
                for (int i = 0; i < tables.Rows.Count; i++)
                {
                    var currentMaxId = ora.SelectMaxId(tables.Rows[i][0].ToString());
                    if (currentMaxId > maxId)
                        maxId = currentMaxId;
                }

                Console.WriteLine(maxId);
            }
            #endregion

            #region 125 getDataForEzhkh
            else if (type == 125)
            {
                var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("1");
                int month = 2;
                int year = 2016;
                var tables = pg.SelectChargeForEzhkh(month, year);
                Int32 row = 1;
                for (int i = 0; i < tables.Rows.Count; i++)
                {
                    ws.Cell(row, 1).Value = tables.Rows[i][0].ToString();
                    ws.Cell(row, 2).Value = tables.Rows[i][1].ToString();
                    ws.Cell(row, 3).Value = tables.Rows[i][2].ToString();
                    ws.Cell(row, 4).Value = tables.Rows[i][3].ToString();
                    ws.Cell(row, 5).Value = tables.Rows[i][4].ToString();
                    ws.Cell(row, 6).Value = tables.Rows[i][5].ToString();
                    ws.Cell(row, 7).Value = tables.Rows[i][6].ToString();
                    ws.Cell(row, 8).Value = tables.Rows[i][7].ToString();
                    ws.Cell(row, 9).Value = tables.Rows[i][8].ToString();
                    ws.Cell(row, 10).Value = tables.Rows[i][9].ToString();
                    ws.Cell(row, 11).Value = tables.Rows[i][10].ToString();
                    ws.Cell(row, 12).Value = tables.Rows[i][11].ToString();
                    ws.Cell(row, 13).Value = tables.Rows[i][12].ToString();
                    ws.Cell(row, 14).Value = tables.Rows[i][13].ToString();
                    ws.Cell(row, 15).Value = tables.Rows[i][14].ToString();
                    row++;
                }

                wb.SaveAs(@"C:\temp\EzhkhImport_"+month+"_"+year+".xlsx");
            }
            #endregion

            #region 126 фаил для Бегина
            else if (type == 126)
            {
                var wb1 = new XLWorkbook(@"C:\Temp\ЖКС.xlsx");
                for (int i = 3; i <= 4826; i++)//5
                {
                    string isFind = wb1.Worksheet(1).Row(i).Cell(9).Value.ToString();
                    if (isFind != "")
                    {
                        Int32 gkhCode = Convert.ToInt32(wb1.Worksheet(1).Row(i).Cell(7).Value.ToString());
                        if (gkhCode == 0)
                        {
                            wb1.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Yellow;
                            wb1.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Yellow;
                            wb1.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Yellow;
                            wb1.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Yellow;
                            wb1.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Yellow;
                        }
                        else if (gkhCode == 1)
                        {
                            wb1.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Green;
                            wb1.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Green;
                            wb1.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Green;
                            wb1.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Green;
                            wb1.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Green;
                        }
                        else
                        {
                            var data = ora.SelectDataToRep(gkhCode);
                            wb1.Worksheet(1).Row(i).Cell(8).Value = data.Rows[0][0] != null ? data.Rows[0][0].ToString() : "";
                            wb1.Worksheet(1).Row(i).Cell(9).Value = data.Rows[0][1] != null ? data.Rows[0][1].ToString() : "";
                            wb1.Worksheet(1).Row(i).Cell(10).Value =
                                Convert.ToInt32(data.Rows[0][2] != null && data.Rows[0][2].ToString() != "" ? data.Rows[0][2].ToString() : "0") +
                                Convert.ToInt32(data.Rows[0][3] != null && data.Rows[0][3].ToString() != "" ? data.Rows[0][3].ToString() : "0");
                            wb1.Worksheet(1).Row(i).Cell(11).Value = data.Rows[0][2] != null && data.Rows[0][2].ToString() != "" ? data.Rows[0][2].ToString() : "0";
                            wb1.Worksheet(1).Row(i).Cell(12).Value = data.Rows[0][3] != null && data.Rows[0][3].ToString() != "" ? data.Rows[0][3].ToString() : "0";
                            wb1.Worksheet(1).Row(i).Cell(13).Value = data.Rows[0][4].ToString();
                            wb1.Worksheet(1).Row(i).Cell(14).Value = data.Rows[0][5].ToString();
                            wb1.Worksheet(1).Row(i).Cell(15).Value = data.Rows[0][6].ToString();
                            wb1.Worksheet(1).Row(i).Cell(16).Value = data.Rows[0][7].ToString();
                        }
                    }
                    else
                    {
                        wb1.Worksheet(1).Row(i).Cell(2).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb1.Worksheet(1).Row(i).Cell(3).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb1.Worksheet(1).Row(i).Cell(4).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb1.Worksheet(1).Row(i).Cell(5).Style.Fill.BackgroundColor = XLColor.Yellow;
                        wb1.Worksheet(1).Row(i).Cell(6).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                    
                }

                wb1.Save();
            }
            #endregion

            #region 127
            else if (type == 127)
            {
                //DateTime dt1 = Convert.ToDateTime("2015-01-01");
                //DateTime dt2 = Convert.ToDateTime("2015-12-31");
                //var nMonth = (int) dt2.Subtract(dt1).TotalDays;
                Decimal d1 = Convert.ToDecimal("0,00");

                Console.WriteLine(d1);
            }
            #endregion

            #region 128 фаил для Бегина
            else if (type == 128)
            {
                string[] stringSeparators = new string[] { "д." };
                int ws = 2;
                var wb1 = new XLWorkbook(@"C:\Temp\zhks(1).xlsx");
                for (int i = 2; i <= 2114; i++)//5
                {
                    if(i%50 == 0)
                        Console.WriteLine(i);
                    string isFind = wb1.Worksheet(1).Row(i).Cell(3).Value.ToString();
                    if (isFind != "")
                    {
                        if (isFind == "")
                        {
                            wb1.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                        }
                    }                 
                    else
                    {
                        if (wb1.Worksheet(1).Row(i).Cell(2).Value.ToString().Trim() == "Невская д.7")
                        {
                            wb1.Worksheet(1).Row(i).Cell(14).Value = "8800801";
                            wb1.Worksheet(1).Row(i).Cell(15).Value = "ул. Невская, д. 7";
                        }
                        else
                        {
                            //string addr = wb1.Worksheet(1).Row(i).Cell(7).Value.ToString().Trim();
                            //string ul = wb1.Worksheet(1).Row(i).Cell(7).Value.ToString().Trim().Split(stringSeparators, StringSplitOptions.None)[0].Trim();
                            //string dom = wb1.Worksheet(1).Row(i).Cell(7).Value.ToString().Trim().Split(stringSeparators, StringSplitOptions.None)[1].Trim();
                            string addr = wb1.Worksheet(1).Row(i).Cell(1).Value.ToString().Trim();
                            string ul = wb1.Worksheet(1).Row(i).Cell(1).Value.ToString().Trim();
                            string dom = wb1.Worksheet(1).Row(i).Cell(1).Value.ToString().Trim().Replace("-","");
                            DataTable gkhCode = pg.SelectGkhCode(ul, dom);
                            if (gkhCode == null)
                            {
                                wb1.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                            }
                            else
                            {
                                wb1.Worksheet(1).Row(i).Cell(2).Value = gkhCode.Rows[0][0].ToString();
                                wb1.Worksheet(1).Row(i).Cell(3).Value = gkhCode.Rows[0][1].ToString();
                                if (gkhCode.Rows[0][0] == null || gkhCode.Rows[0][0].ToString() == "")
                                {
                                    wb1.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Red;
                                }
                            }
                        }
                    }

                }

                wb1.Save();
            }
            #endregion

            #region 129 корректировка по пеням
            else if (type == 129)
            {
                string database = "billAuk";
                string comment = "Выравнивание сальдо";
                var dt = pg.GetMinusPeni();
                //for (int i = 1; i < 1; i++)
                for (int i = 28; i < dt.Rows.Count; i++)             
                {
                    var dtPeniBySupp = pg.GetPeniSuppByNzpKvar(dt.Rows[i][0].ToString());
                    decimal totalPeni = 0;
                    for (int j = 0; j < dtPeniBySupp.Rows.Count; j++)
                    {
                        if (Convert.ToDecimal(dtPeniBySupp.Rows[j][1]) < 0)
                        {
                            int nzp_doc_base = pg.InsertDocBase(database, comment);
                            pg.InsertPerekidka(database,
                                Convert.ToInt32(dtPeniBySupp.Rows[j][0]),
                                Convert.ToDecimal(dtPeniBySupp.Rows[j][1]) * (-1),
                                nzp_doc_base,
                                Convert.ToInt32(dtPeniBySupp.Rows[j][2]),
                                500,
                                Convert.ToInt32(dtPeniBySupp.Rows[j][3]),
                                "2016-03-21", 03, 2016);
                            totalPeni += Convert.ToDecimal(dtPeniBySupp.Rows[j][1]);
                        }
                        
                    }
                    decimal totalRsumTarif = 0;
                    var dtRsumTarifBySuppAndServ = pg.GetRsumTarifSuppAndServByNzpKvar(dt.Rows[i][0].ToString());
                    if (dtRsumTarifBySuppAndServ.Rows.Count == 0)
                    {
                        Console.WriteLine("нет начислений у квартиры: " + dt.Rows[i][0].ToString());
                        var dtFirstServ = pg.GetFirstRsumTarifSuppAndServByNzpKvar(dt.Rows[i][0].ToString());
                        int nzp_doc_base = pg.InsertDocBase(database, comment);
                        pg.InsertPerekidka(database,
                            Convert.ToInt32(dtFirstServ.Rows[0][0]),
                            totalPeni,
                            nzp_doc_base,
                            Convert.ToInt32(dtFirstServ.Rows[0][2]),
                            Convert.ToInt32(dtFirstServ.Rows[0][3]),
                            Convert.ToInt32(dtFirstServ.Rows[0][4]),
                            "2016-03-21", 03, 2016);
                    }
                    else
                    {
                        for (int j = 0; j < dtRsumTarifBySuppAndServ.Rows.Count; j++)
                        {
                            totalRsumTarif += Convert.ToDecimal(dtRsumTarifBySuppAndServ.Rows[j][1]);
                        }
                        decimal koef = totalPeni / totalRsumTarif;
                        decimal writeInBase = 0;
                        for (int j = 0; j < dtRsumTarifBySuppAndServ.Rows.Count; j++)
                        {
                            int nzp_doc_base = pg.InsertDocBase(database, comment);
                            pg.InsertPerekidka(database,
                                Convert.ToInt32(dtRsumTarifBySuppAndServ.Rows[j][0]),
                                Math.Round(Convert.ToDecimal(dtRsumTarifBySuppAndServ.Rows[j][1]) * koef, 2),
                                nzp_doc_base,
                                Convert.ToInt32(dtRsumTarifBySuppAndServ.Rows[j][2]),
                                Convert.ToInt32(dtRsumTarifBySuppAndServ.Rows[j][3]),
                                Convert.ToInt32(dtRsumTarifBySuppAndServ.Rows[j][4]),
                                "2016-03-21", 03, 2016);
                            writeInBase += Math.Round(Convert.ToDecimal(dtRsumTarifBySuppAndServ.Rows[j][1]) * koef, 2);
                        }
                        if (writeInBase != totalPeni)
                        {
                            int nzp_doc_base = pg.InsertDocBase(database, comment);
                            pg.InsertPerekidka(database,
                                Convert.ToInt32(dtRsumTarifBySuppAndServ.Rows[0][0]),
                                totalPeni - writeInBase,
                                nzp_doc_base,
                                Convert.ToInt32(dtRsumTarifBySuppAndServ.Rows[0][2]),
                                Convert.ToInt32(dtRsumTarifBySuppAndServ.Rows[0][3]),
                                Convert.ToInt32(dtRsumTarifBySuppAndServ.Rows[0][4]),
                                "2016-03-21", 03, 2016);
                        }
                    }
                    
                }
            }
            #endregion

            #region 130 Складываем отчет по начислению и оплате за месяца
            else if (type == 130)
            {
                var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("1");
                List<string> months = new List<string>();
                months.Add("сентябрь");
                months.Add("октябрь");
                months.Add("ноябрь");
                months.Add("декабрь");
                String address = "";
                String currServ = "";
                List<string> addedAddress = new List<string>();
                List<string> addedServ = new List<string>();
                List<AddressServ> addressServ = new List<AddressServ>();
                foreach (string month in months)
                {
                    var wb1 = new XLWorkbook(@"C:\Temp\" + month + ".xlsx");
                    for (int i = 2; i <= 110; i++)
                    {
                        if (wb1.Worksheet(1).Row(i).Cell(1).Value.ToString() == "")
                            break;

                        if (wb1.Worksheet(1).Row(i).Cell(1).Value.ToString().Contains("ТОЛЬЯТТИ Г") &&
                            address != wb1.Worksheet(1).Row(i).Cell(1).Value.ToString().Trim())
                        {
                            address = wb1.Worksheet(1).Row(i).Cell(1).Value.ToString().Trim();
                            i = i + 3;
                        }
                            
                        else if (wb1.Worksheet(1).Row(i).Cell(1).Value.ToString().Contains("Сводная по домам") &&
                                 address != wb1.Worksheet(1).Row(i).Cell(1).Value.ToString().Trim())
                        {
                            address = wb1.Worksheet(1).Row(i).Cell(1).Value.ToString().Trim();
                            i = i + 3;
                        }
                            

                        if (!addedAddress.Contains(address))
                        {                           
                            List<Service> servic = new List<Service>();
                            AddressServ addrServ = new AddressServ();
                            addrServ.Address = address;
                            addrServ.Services = servic;

                            Service serv = new Service();
                            string str = wb1.Worksheet(1).Row(i).Cell(5).Value.ToString().Trim();
                            serv.Nedop = Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(5).Value.ToString().Trim().Replace('.',','));
                            serv.Peni = Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(7).Value.ToString().Trim().Replace('.', ','));
                            serv.Reval = Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(4).Value.ToString().Trim().Replace('.', ','));
                            serv.RsumTarif = Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(3).Value.ToString().Trim().Replace('.', ','));
                            serv.Serv = wb1.Worksheet(1).Row(i).Cell(1).Value.ToString().Trim();
                            serv.SumInsaldo = Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(2).Value.ToString().Trim().Replace('.', ','));
                            serv.SumMoney = Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(8).Value.ToString().Trim().Replace('.', ','));
                            serv.SumMoneyPeni = Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(9).Value.ToString().Trim().Replace('.', ','));
                            serv.SumOutsaldo = Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(10).Value.ToString().Trim().Replace('.', ','));
                            serv.TotalRsum = Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(6).Value.ToString().Trim().Replace('.', ','));
                            addrServ.Services.Add(serv);

                            addressServ.Add(addrServ);

                            addedAddress.Add(address);
                            addedServ.Add(address + "|" + wb1.Worksheet(1).Row(i).Cell(1).Value.ToString().Trim());
                        }
                        else
                        {
                            currServ = wb1.Worksheet(1).Row(i).Cell(1).Value.ToString().Trim();
                            if (!addedServ.Contains(address + "|" + currServ))
                            {
                                Service serv = new Service();
                                serv.Nedop = Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(5).Value.ToString().Trim().Replace('.', ','));
                                serv.Peni = Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(7).Value.ToString().Trim().Replace('.', ','));
                                serv.Reval = Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(4).Value.ToString().Trim().Replace('.', ','));
                                serv.RsumTarif = Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(3).Value.ToString().Trim().Replace('.', ','));
                                serv.Serv = wb1.Worksheet(1).Row(i).Cell(1).Value.ToString().Trim();
                                serv.SumInsaldo = Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(2).Value.ToString().Trim().Replace('.', ','));
                                serv.SumMoney = Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(8).Value.ToString().Trim().Replace('.', ','));
                                serv.SumMoneyPeni = Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(9).Value.ToString().Trim().Replace('.', ','));
                                serv.SumOutsaldo = Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(10).Value.ToString().Trim().Replace('.', ','));
                                serv.TotalRsum = Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(6).Value.ToString().Trim().Replace('.', ','));
                                switch (address)
                                {
                                    case "ТОЛЬЯТТИ Г, -,РЕВОЛЮЦИОННАЯ д.40":
                                    {
                                        addressServ[0].Services.Add(serv);
                                        break;
                                    }
                                    case "ТОЛЬЯТТИ Г, -,РЕВОЛЮЦИОННАЯ д.44":
                                    {
                                        addressServ[1].Services.Add(serv);
                                        break;
                                    }
                                    case "ТОЛЬЯТТИ Г, -,РЕВОЛЮЦИОННАЯ д.50":
                                    {
                                        addressServ[2].Services.Add(serv);
                                        break;
                                    }
                                    case "Сводная по домам":
                                    {
                                        addressServ[3].Services.Add(serv);
                                        break;
                                    }
                                }
                                addedServ.Add(address + "|" + wb1.Worksheet(1).Row(i).Cell(1).Value.ToString().Trim());
                            }
                            else
                            {
                                switch (address)
                                {
                                    case "ТОЛЬЯТТИ Г, -,РЕВОЛЮЦИОННАЯ д.40":
                                    {
                                        List<Service> currSrv = addressServ[0].Services;
                                        Service newServ = new Service();
                                        //Интерфейс.Сообщение.Показать("Район= " + текущийРайон + "; Название= "+ текущаяКомпания + "; КУ= " + текущаяКУ);
                                        newServ = currSrv.Find(
                                            delegate(Service s)
                                            {
                                                return s.Serv == currServ;
                                            }
                                        );
                                        currSrv.Remove(newServ);
                                        newServ.Nedop += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(5).Value.ToString().Trim().Replace('.', ','));
                                        newServ.Peni += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(7).Value.ToString().Trim().Replace('.', ','));
                                        newServ.Reval += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(4).Value.ToString().Trim().Replace('.', ','));
                                        newServ.RsumTarif += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(3).Value.ToString().Trim().Replace('.', ','));
                                        newServ.SumInsaldo += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(2).Value.ToString().Trim().Replace('.', ','));
                                        newServ.SumMoney += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(8).Value.ToString().Trim().Replace('.', ','));
                                        newServ.SumMoneyPeni += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(9).Value.ToString().Trim().Replace('.', ','));
                                        newServ.SumOutsaldo += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(10).Value.ToString().Trim().Replace('.', ','));
                                        newServ.TotalRsum += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(6).Value.ToString().Trim().Replace('.', ','));
                                        currSrv.Add(newServ);
                                        addressServ[0].Services = currSrv;
                                        break;
                                    }
                                    case "ТОЛЬЯТТИ Г, -,РЕВОЛЮЦИОННАЯ д.44":
                                        {
                                            List<Service> currSrv = addressServ[1].Services;
                                            Service newServ = new Service();
                                            //Интерфейс.Сообщение.Показать("Район= " + текущийРайон + "; Название= "+ текущаяКомпания + "; КУ= " + текущаяКУ);
                                            newServ = currSrv.Find(
                                                delegate(Service s)
                                                {
                                                    return s.Serv == currServ;
                                                }
                                            );
                                            currSrv.Remove(newServ);
                                            newServ.Nedop += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(5).Value.ToString().Trim().Replace('.', ','));
                                            newServ.Peni += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(7).Value.ToString().Trim().Replace('.', ','));
                                            newServ.Reval += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(4).Value.ToString().Trim().Replace('.', ','));
                                            newServ.RsumTarif += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(3).Value.ToString().Trim().Replace('.', ','));
                                            newServ.SumInsaldo += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(2).Value.ToString().Trim().Replace('.', ','));
                                            newServ.SumMoney += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(8).Value.ToString().Trim().Replace('.', ','));
                                            newServ.SumMoneyPeni += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(9).Value.ToString().Trim().Replace('.', ','));
                                            newServ.SumOutsaldo += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(10).Value.ToString().Trim().Replace('.', ','));
                                            newServ.TotalRsum += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(6).Value.ToString().Trim().Replace('.', ','));
                                            currSrv.Add(newServ);
                                            addressServ[1].Services = currSrv;
                                            break;
                                        }
                                    case "ТОЛЬЯТТИ Г, -,РЕВОЛЮЦИОННАЯ д.50":
                                        {
                                            List<Service> currSrv = addressServ[2].Services;
                                            Service newServ = new Service();
                                            //Интерфейс.Сообщение.Показать("Район= " + текущийРайон + "; Название= "+ текущаяКомпания + "; КУ= " + текущаяКУ);
                                            newServ = currSrv.Find(
                                                delegate(Service s)
                                                {
                                                    return s.Serv == currServ;
                                                }
                                            );
                                            currSrv.Remove(newServ);
                                            newServ.Nedop += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(5).Value.ToString().Trim().Replace('.', ','));
                                            newServ.Peni += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(7).Value.ToString().Trim().Replace('.', ','));
                                            newServ.Reval += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(4).Value.ToString().Trim().Replace('.', ','));
                                            newServ.RsumTarif += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(3).Value.ToString().Trim().Replace('.', ','));
                                            newServ.SumInsaldo += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(2).Value.ToString().Trim().Replace('.', ','));
                                            newServ.SumMoney += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(8).Value.ToString().Trim().Replace('.', ','));
                                            newServ.SumMoneyPeni += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(9).Value.ToString().Trim().Replace('.', ','));
                                            newServ.SumOutsaldo += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(10).Value.ToString().Trim().Replace('.', ','));
                                            newServ.TotalRsum += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(6).Value.ToString().Trim().Replace('.', ','));
                                            currSrv.Add(newServ);
                                            addressServ[2].Services = currSrv;
                                            break;
                                        }
                                    case "Сводная по домам":
                                        {
                                            List<Service> currSrv = addressServ[3].Services;
                                            Service newServ = new Service();
                                            //Интерфейс.Сообщение.Показать("Район= " + текущийРайон + "; Название= "+ текущаяКомпания + "; КУ= " + текущаяКУ);
                                            newServ = currSrv.Find(
                                                delegate(Service s)
                                                {
                                                    return s.Serv == currServ;
                                                }
                                            );
                                            currSrv.Remove(newServ);
                                            newServ.Nedop += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(5).Value.ToString().Trim().Replace('.', ','));
                                            newServ.Peni += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(7).Value.ToString().Trim().Replace('.', ','));
                                            newServ.Reval += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(4).Value.ToString().Trim().Replace('.', ','));
                                            newServ.RsumTarif += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(3).Value.ToString().Trim().Replace('.', ','));
                                            newServ.SumInsaldo += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(2).Value.ToString().Trim().Replace('.', ','));
                                            newServ.SumMoney += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(8).Value.ToString().Trim().Replace('.', ','));
                                            newServ.SumMoneyPeni += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(9).Value.ToString().Trim().Replace('.', ','));
                                            newServ.SumOutsaldo += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(10).Value.ToString().Trim().Replace('.', ','));
                                            newServ.TotalRsum += Convert.ToDecimal(wb1.Worksheet(1).Row(i).Cell(6).Value.ToString().Trim().Replace('.', ','));
                                            currSrv.Add(newServ);
                                            addressServ[3].Services = currSrv;
                                            break;
                                        }
                                }
                            }
                        }
                    }
                }
                List<String> headers = new List<string>();
                headers.Add("Услуга");
                headers.Add("Сальдо на начало периода");
                headers.Add("Начислено 100%");
                headers.Add("Перерасчеты");
                headers.Add("Недопоставки");
                headers.Add("Итого начислено");
                headers.Add("Начислено пени");
                headers.Add("Оплачено");
                headers.Add("Оплачено пени");
                headers.Add("Сальдо на конец периода");
                Int32 row = 2;
                foreach (AddressServ addrServ in addressServ)
                {
                    ws.Cell(row, 1).Value = addrServ.Address;
                    row++;
                    Int32 col = 1;
                    foreach (String header in headers)
                    {
                        ws.Cell(row, col).Value = header;
                        col++;
                    }
                    row++;
                    for (int i = 1; i <= 10; i++)
                    {
                        ws.Cell(row, i).Value = i;
                    }
                    row++;
                    foreach (Service serv in addrServ.Services)
                    {
                        ws.Cell(row, 1).Value = serv.Serv;
                        ws.Cell(row, 2).Value = serv.SumInsaldo;
                        ws.Cell(row, 3).Value = serv.RsumTarif;
                        ws.Cell(row, 4).Value = serv.Reval;
                        ws.Cell(row, 5).Value = serv.Nedop;
                        ws.Cell(row, 6).Value = serv.TotalRsum;
                        ws.Cell(row, 7).Value = serv.Peni;
                        ws.Cell(row, 8).Value = serv.SumMoney;
                        ws.Cell(row, 9).Value = serv.SumMoneyPeni;
                        ws.Cell(row, 10).Value = serv.SumOutsaldo;
                        row++;
                    }
                }

                wb.SaveAs(@"C:\temp\ReporSumChargePrih.xlsx");
            }
            #endregion

            #region 131 Тлт выгрузка
            else if (type == 131)
            {
                var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("Лист1");
                var tableSum = pg.GetSaldoAndParam();
                int row = 2;
                for(int i = 0; i<tableSum.Rows.Count; i++)
                {               
                    ws.Cell(row, 1).Value = "Тольятти";
                    ws.Cell(row, 2).Value = tableSum.Rows[i][0].ToString();
                    ws.Cell(row, 3).Value = tableSum.Rows[i][1].ToString();
                    ws.Cell(row, 4).Value = tableSum.Rows[i][2].ToString();
                    ws.Cell(row, 5).Value = tableSum.Rows[i][3].ToString();
                    ws.Cell(row, 6).Value = tableSum.Rows[i][4].ToString();
                    ws.Cell(row, 7).Value = tableSum.Rows[i][5].ToString();
                    ws.Cell(row, 8).Value = tableSum.Rows[i][6].ToString();
                    ws.Cell(row, 9).Value = tableSum.Rows[i][7].ToString();
                    ws.Cell(row, 10).Value = tableSum.Rows[i][8].ToString();
                    ws.Cell(row, 11).Value = tableSum.Rows[i][9].ToString();

                    string[] objsObj = tableSum.Rows[i][10].ToString().Split(new[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                    if (objsObj.Count() >= 3)
                    {
                        ws.Cell(row, 12).Value = objsObj[0];
                        ws.Cell(row, 13).Value = objsObj[1];
                        string other = "";
                        for (int j = 2; j < objsObj.Count(); j++)
                        {
                            other += objsObj[j] + " ";
                        }
                        ws.Cell(row, 14).Value = other.Trim();
                    }
                    ws.Cell(row, 15).Value = tableSum.Rows[i][11].ToString();
                    ws.Cell(row, 16).Value = tableSum.Rows[i][12].ToString();
                    ws.Cell(row, 17).Value = tableSum.Rows[i][13].ToString();
                    ws.Cell(row, 18).Value = tableSum.Rows[i][14].ToString();
                    row++;
                }

                var ws2 = wb.Worksheets.Add("Лист2");
                var tableByServ = pg.GetSaldoAndParamByServ();
                row = 2;
                for (int i = 0; i < tableByServ.Rows.Count; i++)
                {
                    ws2.Cell(row, 1).Value = "Тольятти";
                    ws2.Cell(row, 2).Value = tableByServ.Rows[i][0].ToString();
                    ws2.Cell(row, 3).Value = tableByServ.Rows[i][1].ToString();
                    ws2.Cell(row, 4).Value = tableByServ.Rows[i][2].ToString();
                    ws2.Cell(row, 5).Value = tableByServ.Rows[i][3].ToString();
                    ws2.Cell(row, 6).Value = tableByServ.Rows[i][4].ToString();
                    ws2.Cell(row, 7).Value = tableByServ.Rows[i][5].ToString();
                    ws2.Cell(row, 8).Value = tableByServ.Rows[i][6].ToString();
                    ws2.Cell(row, 9).Value = tableByServ.Rows[i][7].ToString();
                    ws2.Cell(row, 10).Value = tableByServ.Rows[i][8].ToString();
                    ws2.Cell(row, 11).Value = tableByServ.Rows[i][9].ToString();

                    string[] objsObj = tableByServ.Rows[i][10].ToString().Split(new[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                    if (objsObj.Count() >= 3)
                    {
                        ws2.Cell(row, 12).Value = objsObj[0];
                        ws2.Cell(row, 13).Value = objsObj[1];
                        string other = "";
                        for (int j = 2; j < objsObj.Count(); j++)
                        {
                            other += objsObj[j] + " ";
                        }
                        ws2.Cell(row, 14).Value = other.Trim();
                    }
                    ws2.Cell(row, 15).Value = tableByServ.Rows[i][11].ToString();
                    ws2.Cell(row, 16).Value = tableByServ.Rows[i][12].ToString();
                    ws2.Cell(row, 17).Value = tableByServ.Rows[i][13].ToString();
                    ws2.Cell(row, 18).Value = tableByServ.Rows[i][14].ToString();
                    row++;
                }

                wb.SaveAs(@"C:\temp\TltOut.xlsx");
            }
            #endregion

            #region 132 Тлт выгрузка счетчики
            else if (type == 132)
            {
                var wb = new XLWorkbook();
                List<string> months = new List<string>();
                months.Add("09.2015");
                months.Add("10.2015");
                months.Add("11.2015");
                months.Add("12.2015");
                months.Add("01.2016");
                months.Add("02.2016");
                months.Add("03.2016");
                months.Add("04.2016");
                foreach (string month in months)
                {
                    var ws = wb.Worksheets.Add("Лист_" + month);
                    var tableSum = pg.GetCountersVal(month);
                    int row = 2;
                    for (int i = 0; i < tableSum.Rows.Count; i++)
                    {
                        ws.Cell(row, 1).Value = "Тольятти";
                        ws.Cell(row, 2).Value = tableSum.Rows[i][0].ToString();
                        ws.Cell(row, 3).Value = tableSum.Rows[i][1].ToString();
                        ws.Cell(row, 4).Value = tableSum.Rows[i][2].ToString();
                        ws.Cell(row, 5).Value = tableSum.Rows[i][3].ToString();
                        ws.Cell(row, 6).Value = tableSum.Rows[i][4].ToString();
                        ws.Cell(row, 7).Value = tableSum.Rows[i][5].ToString();
                        ws.Cell(row, 8).Value = tableSum.Rows[i][6].ToString();
                        ws.Cell(row, 9).Value = tableSum.Rows[i][7].ToString();
                        ws.Cell(row, 10).Value = tableSum.Rows[i][8].ToString();
                        ws.Cell(row, 11).Value = tableSum.Rows[i][9].ToString();
                        ws.Cell(row, 12).Value = tableSum.Rows[i][10].ToString();
                        ws.Cell(row, 13).Value = tableSum.Rows[i][11].ToString();
                        row++;
                    }
                }
                

                wb.SaveAs(@"C:\temp\TltOutCounters.xlsx");
            }
            #endregion

            #region 133 Тлт выгрузка домофон
            else if (type == 133)
            {
                var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("Лист1");
                var tableSum = pg.GetTarifDomofon();
                int row = 2;
                for (int i = 0; i < tableSum.Rows.Count; i++)
                {
                    ws.Cell(row, 1).Value = "Тольятти";
                    ws.Cell(row, 2).Value = tableSum.Rows[i][0].ToString();
                    ws.Cell(row, 3).Value = tableSum.Rows[i][1].ToString();
                    ws.Cell(row, 4).Value = tableSum.Rows[i][2].ToString();
                    ws.Cell(row, 5).Value = tableSum.Rows[i][3].ToString();
                    ws.Cell(row, 6).Value = tableSum.Rows[i][4].ToString();
                    ws.Cell(row, 7).Value = "Домофон";
                    ws.Cell(row, 8).Value = tableSum.Rows[i][5].ToString();
                    row++;
                }
                wb.SaveAs(@"C:\temp\TltOutDomofon.xlsx");
            }
            #endregion

            #region 134 Тлт выгрузка перечень услуг
            else if (type == 134)
            {
                var wb = new XLWorkbook();
                var ws = wb.Worksheets.Add("Лист1");
                var tableSum = pg.GetServList();
                int row = 2;
                for (int i = 0; i < tableSum.Rows.Count; i++)
                {
                    ws.Cell(row, 1).Value = "Тольятти";
                    ws.Cell(row, 2).Value = tableSum.Rows[i][0].ToString();
                    ws.Cell(row, 3).Value = tableSum.Rows[i][1].ToString();
                    ws.Cell(row, 4).Value = tableSum.Rows[i][2].ToString();
                    ws.Cell(row, 5).Value = tableSum.Rows[i][3].ToString();
                    ws.Cell(row, 6).Value = tableSum.Rows[i][4].ToString();
                    ws.Cell(row, 7).Value = tableSum.Rows[i][5].ToString();
                    ws.Cell(row, 8).Value = tableSum.Rows[i][6].ToString();
                    ws.Cell(row, 9).Value = tableSum.Rows[i][7].ToString();
                    ws.Cell(row, 10).Value = tableSum.Rows[i][8].ToString();
                    row++;
                }
                wb.SaveAs(@"C:\temp\TltOutServList.xlsx");
            }
            #endregion

            #region 995
            else if (type == 995)
            {
                decimal d = 7.5308914975270384m;
                Console.WriteLine(Math.Round(d).ToString());
                List<string> str = new List<string>();
            }
            #endregion

            #region 996
            else if (type == 996)
            {

                DateTime d = new DateTime(2015, 4, 22, 12, 0, 1);
                DateTime _scheduleTime = d.AddMinutes(40);
                Console.WriteLine(_scheduleTime);
                double interval = _scheduleTime.Subtract(d).TotalSeconds * 1000;
                Console.WriteLine(interval / 1000);
            }
            #endregion

            #region 997
            else if (type == 997)
            {
                string url = "http://localhost:56243/request/UpdateLot";
                WebRequest request = WebRequest.Create(url);
                request.Timeout = 60 * 60 * 1000;
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            }
            #endregion

            #region 998
            else if (type == 998)
            {
                string s = "fcsProtocolOK2_0142200001314016749_3318261";
                Console.WriteLine(s.Split('_')[0].Substring(s.Split('_')[0].Length - 1));
            }
            #endregion

            #region 999
            else if (type == 999)
            {
                Random rnd = new Random();
                Console.Write("Введите стартовое число: ");
                int start = Convert.ToInt32(Console.ReadLine());
                Console.Write("Введите конечное число: ");
                int finish = Convert.ToInt32(Console.ReadLine());
                int num = rnd.Next(start, finish); // creates a number between 1 and 12
                Console.WriteLine(num);
            }
            #endregion
            #region 1000
            else if (type == 1000)
            {
                string str = "17 БЛОК Б";
                int pos = 0;
                int isNum;
                foreach (char c in str)
                {
                    bool result = Int32.TryParse(c.ToString(), out isNum);
                    if (result)
                        pos++;
                    else
                        break;
                }
                Console.WriteLine(str.Substring(0, pos).Trim());
                Console.WriteLine(str.Substring(pos).Trim());
            }
            #endregion

            #region 1001
            else if (type == 1001)
            {
               ora.TestHome();
            }
            #endregion
            Console.WriteLine("Готово!!!");
            Console.ReadLine();
        }

        public class Service
        {
            public String Serv { get; set; }
            public Decimal SumInsaldo { get; set; }
            public Decimal RsumTarif { get; set; }
            public Decimal Reval { get; set; }
            public Decimal Nedop { get; set; }
            public Decimal TotalRsum { get; set; }
            public Decimal Peni { get; set; }
            public Decimal SumMoney { get; set; }
            public Decimal SumMoneyPeni { get; set; }
            public Decimal SumOutsaldo { get; set; }
        }

        public class AddressServ
        {
            public String Address { get; set; }
            public List<Service> Services { get; set; }
        }
    }   
}



/*if (dtHouse.Rows[i][2] != null && dtHouse.Rows[i][2].ToString() != "")
                            {
                                for (int j = 0; j < dtNorm.Rows.Count; j++)
                                {
                                    if (dtHouse.Rows[i][1].ToString() == dtNorm.Rows[j][0].ToString() && dtHouse.Rows[i][2].ToString() == dtNorm.Rows[j][2].ToString())
                                    {
                                        ws.Cell(rows, 1).Value = dtHouse.Rows[i][0].ToString();
                                        string str = "";
                                        str = Convert2(dtNorm.Rows[j][1].ToString(), Encoding.GetEncoding(1251), Encoding.Default);
                                        ws.Cell(rows, 2).Value = str;
                                        //ws.Cell(i + 1, 2).Value = dtNormTemp.Rows[i][1].ToString();
                                        ws.Cell(rows, 3).Value = dtNorm.Rows[j][3].ToString();
                                        kHouse.Add(dtHouse.Rows[i][0].ToString());
                                        rows++;
                                        break;
                                    }
                                }
                            }
                            else if (dtHouse.Rows[i][3] != null && dtHouse.Rows[i][3].ToString() != "")
                            {
                                for (int j = 0; j < dtNorm.Rows.Count; j++)
                                {
                                    if (dtHouse.Rows[i][1].ToString() == dtNorm.Rows[j][0].ToString() && dtHouse.Rows[i][3].ToString() == dtNorm.Rows[j][2].ToString())
                                    {
                                        ws.Cell(rows, 1).Value = dtHouse.Rows[i][0].ToString();
                                        string str = "";
                                        str = Convert2(dtNorm.Rows[j][1].ToString(), Encoding.GetEncoding(1251), Encoding.Default);
                                        ws.Cell(rows, 2).Value = str;
                                        //ws.Cell(i + 1, 2).Value = dtNormTemp.Rows[i][1].ToString();
                                        ws.Cell(rows, 3).Value = dtNorm.Rows[j][3].ToString();
                                        kHouse.Add(dtHouse.Rows[i][0].ToString());
                                        rows++;
                                        break;
                                    }
                                }
                            }
                            else if (dtHouse.Rows[i][4] != null && dtHouse.Rows[i][4].ToString() != "")
                            {
                                for (int j = 0; j < dtNorm.Rows.Count; j++)
                                {
                                    if (dtHouse.Rows[i][1].ToString() == dtNorm.Rows[j][0].ToString() && dtHouse.Rows[i][4].ToString() == dtNorm.Rows[j][2].ToString())
                                    {
                                        ws.Cell(rows, 1).Value = dtHouse.Rows[i][0].ToString();
                                        string str = "";
                                        str = Convert2(dtNorm.Rows[j][1].ToString(), Encoding.GetEncoding(1251), Encoding.Default);
                                        ws.Cell(rows, 2).Value = str;
                                        //ws.Cell(i + 1, 2).Value = dtNormTemp.Rows[i][1].ToString();
                                        ws.Cell(rows, 3).Value = dtNorm.Rows[j][3].ToString();
                                        kHouse.Add(dtHouse.Rows[i][0].ToString());
                                        rows++;
                                        break;
                                    }
                                }
                            }*/