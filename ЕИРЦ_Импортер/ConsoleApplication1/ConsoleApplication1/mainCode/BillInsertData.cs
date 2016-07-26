using ClosedXML.Excel;
using ConsoleApplication1.Database;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1.mainCode
{
    class BillInsertData
    {
        private pg pg;
        public BillInsertData()
        {
            pg = new pg();
        }

        public void GroupHouseConstruct()
        {
            Start:
            Console.Write("Введите наименование БД:");
            string database = Console.ReadLine();
            Console.Write("Введите наименование параметра:");
            string prm_name = Console.ReadLine();
            string curFile = @"C:\temp\Exp\" + prm_name + ".xlsx";
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
                            if (kvars != null)
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

        public void InsertNormativ2()
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

        public void InsertNormativ()
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


        public void InsertCountersVal()
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

        public void CheckOrInsertCountersVal()
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

        public void InsertCountersVal2()
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

        public void UpdateCountersDatClose()
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

        public void InsertCountersVal3()
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
                        kvarParams = pg.SelectKvarParams("", database, address.Split(separator1, StringSplitOptions.None)[1].Trim(), 7155107);
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
                        int nzp_counter = pg.SelectNzpCounter("", database, kvarParams[0], nzp_serv, Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim(), "", "");
                        if (nzp_counter == 0)
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

        public void ImportKvarFromExcel()
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
                if (res != "Success")
                    Console.WriteLine(res + "|||" + nkvar);
            }
            wb2.Save();
        }

        public void UpdatePackDate()
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

        public void InsertTarif()
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

        public void UpdateOutSaldo()
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

        public void UpdateKvarTotalSquare()
        {
            List<string> prefs = new List<string>();
            prefs.Add("bill01");
            Console.Write("Введите наименование БД:");
            string database = Console.ReadLine();
            var wbPrmName = new XLWorkbook(@"C:\temp\kommunal итог ОКОНЧАТЕЛЬНЫЙ (именно этот файл загружаем на АУК).xlsx");
            BillInsertDataDb billInsertDataDb = new BillInsertDataDb(database);
            for (int i = 2; i <= 1293; i++)
            {
                if (Convert.ToString(wbPrmName.Worksheet(1).Row(i).Cell(10).Value) != "")
                {
                    List<string> nzp = pg.SelectNzpKvarByPkod(Convert.ToString(wbPrmName.Worksheet(1).Row(i).Cell(5).Value).Trim(), database);
                    if(nzp != null)
                        billInsertDataDb.UpdateKvarTotalSquare(Convert.ToString(wbPrmName.Worksheet(1).Row(i).Cell(10).Value).Trim(), nzp[0]);
                }
            }
            wbPrmName.Save();

        }
    }
}
