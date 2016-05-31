using ClosedXML.Excel;
using ConsoleApplication1.Database;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1.mainCode
{
    class BillKart
    {
        public void LoadKart()
        {
            BillKartDb billKartDb = new BillKartDb();

            Console.Write("Введите наименование базы:");
            string database = Console.ReadLine();
            var book = new XLWorkbook(@"C:\temp\часть 1 и 2.xlsx");
            string nzp_kvar = "";
            bool isClear = false;
            for (int i = 2; i <= 11440; i++)
            {
                if (Convert.ToString(book.Worksheet(1).Row(i).Cell(32).Value).Trim() != "")
                    continue;
                if (i % 100 == 0)
                    Console.WriteLine(i);
                if (Convert.ToString(book.Worksheet(1).Row(i).Cell(2).Value).Trim() != "")
                {
                    nzp_kvar = billKartDb.SelectNzpKvar(database,
                       Convert.ToString(book.Worksheet(1).Row(i).Cell(2).Value).Trim(),
                       Convert.ToString(book.Worksheet(1).Row(i).Cell(3).Value).Trim(),
                       Convert.ToString(book.Worksheet(1).Row(i).Cell(4).Value).Trim());
                    isClear = false;
                    continue;
                }

                if (nzp_kvar.Split('|')[0] == "0")
                {
                    book.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                    book.Worksheet(1).Row(i).Cell(31).Value = nzp_kvar.Split('|')[1];
                }
                else if (nzp_kvar.Split('|')[0] == "-1")
                {
                    book.Worksheet(1).Row(i).Style.Fill.BackgroundColor = XLColor.Orange;
                    book.Worksheet(1).Row(i).Cell(31).Value = nzp_kvar.Split('|')[1];
                }
                else
                {
                    if (!isClear)
                    {
                        billKartDb.ClearKart(database, nzp_kvar.Split('|')[0]);
                        isClear = true;
                    }
                    
                    book.Worksheet(1).Row(i).Cell(32).Value = "1";
                    int nzp_gil = billKartDb.InsertGil(database);
                    int nzp_rod = 0;
                    #region nzp_rod
                    switch (Convert.ToString(book.Worksheet(1).Row(i).Cell(100).Value).Trim())
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
                    int tempDoc = 0;
                    if (Convert.ToString(book.Worksheet(1).Row(i).Cell(15).Value).Trim() != "")
                    {
                        bool result =
                            Int32.TryParse(Convert.ToString(book.Worksheet(1).Row(i).Cell(15).Value).Trim(),
                                out tempDoc);
                        if (!result)
                        {
                            nzp_dok = 2;
                        }
                        else
                        {
                            nzp_dok = 10;
                        }
                    }
                    else
                    {
                        nzp_dok = -1;
                    }
                    /*
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
                    */
                    string serij = "";
                    if (Convert.ToString(book.Worksheet(1).Row(i).Cell(15).Value).Trim() != "" && Convert.ToString(book.Worksheet(1).Row(i).Cell(15).Value).Trim().Length >= 3)
                    {
                        if (nzp_dok == 10 &&
                            Convert.ToString(book.Worksheet(1).Row(i).Cell(15).Value).Trim().Length >= 4)
                            serij =
                                Convert.ToString(book.Worksheet(1).Row(i).Cell(15).Value).Trim().Substring(0, 2) +
                                " " +
                                Convert.ToString(book.Worksheet(1).Row(i).Cell(15).Value).Trim().Substring(2, 2);
                        else
                            serij = Convert.ToString(book.Worksheet(1).Row(i).Cell(15).Value).Trim();
                    }

                    string rem_ku = (Convert.ToString(book.Worksheet(1).Row(i).Cell(27).Value).Trim() != ""
                        ? Convert.ToString(book.Worksheet(1).Row(i).Cell(27).Value).Trim() + ", "
                        : "") +
                        (Convert.ToString(book.Worksheet(1).Row(i).Cell(28).Value).Trim() != ""
                        ? Convert.ToString(book.Worksheet(1).Row(i).Cell(28).Value).Trim() + ", "
                        : "") +
                          Convert.ToString(book.Worksheet(1).Row(i).Cell(29).Value).Trim() +
                          (Convert.ToString(book.Worksheet(1).Row(i).Cell(30).Value).Trim() != ""
                            ? ", " + Convert.ToString(book.Worksheet(1).Row(i).Cell(30).Value).Trim()
                            : "");
                    
                    int nzp_kart = billKartDb.InsertKart(database, nzp_gil, nzp_kvar.Split('|')[0],
                                                Convert.ToString(book.Worksheet(1).Row(i).Cell(6).Value).Trim().ToUpper(),
                                                Convert.ToString(book.Worksheet(1).Row(i).Cell(7).Value).Trim().ToUpper(),
                                                Convert.ToString(book.Worksheet(1).Row(i).Cell(8).Value).Trim().ToUpper(),
                                                Convert.ToString(book.Worksheet(1).Row(i).Cell(14).Value).Trim(),
                                                "",
                                                nzp_dok,
                                                serij,
                                                Convert.ToString(book.Worksheet(1).Row(i).Cell(18).Value).Trim(),
                                                Convert.ToString(book.Worksheet(1).Row(i).Cell(17).Value).Trim(),
                                                Convert.ToString(book.Worksheet(1).Row(i).Cell(16).Value).Trim(),
                                                "П",
                                                Convert.ToString(book.Worksheet(1).Row(i).Cell(13).Value).Trim(),
                                                Convert.ToString(book.Worksheet(1).Row(i).Cell(12).Value).Trim(),
                                                nzp_rod,
                                                "",
                                                Convert.ToString(book.Worksheet(1).Row(i).Cell(20).Value).Trim(),
                                                Convert.ToString(book.Worksheet(1).Row(i).Cell(21).Value).Trim(),
                                                Convert.ToString(book.Worksheet(1).Row(i).Cell(22).Value).Trim(),
                                                Convert.ToString(book.Worksheet(1).Row(i).Cell(23).Value).Trim(),
                                                Convert.ToString(book.Worksheet(1).Row(i).Cell(24).Value).Trim(),
                                                Convert.ToString(book.Worksheet(1).Row(i).Cell(25).Value).Trim(),
                                                Convert.ToString(book.Worksheet(1).Row(i).Cell(26).Value).Trim(),
                                                rem_ku);
                    billKartDb.InsertGrgd(nzp_kart);
                }
            }
            book.Save();
        }

        public void LoadKart2()
        {
            BillKartDb billKartDb = new BillKartDb();
            BillBaseDb billBaseDb = new BillBaseDb();

            Console.Write("Введите наименование БД:");
            string database = Console.ReadLine();
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
                    nzp_kvar = billBaseDb.SelectNzpKvarByKvarDom("billAuk",
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
                    int nzp_gil = billKartDb.InsertGil(database);
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
                        int nzp_kart = billKartDb.InsertKart("billAuk", nzp_gil, nzp_kvar.Split('|')[0],
                                                Convert.ToString(book.Worksheet(1).Row(i).Cell(5).Value).Trim().ToUpper().Split(' ')[0],
                                                Convert.ToString(book.Worksheet(1).Row(i).Cell(5).Value).Trim().ToUpper().Split(' ')[1],
                                                Convert.ToString(book.Worksheet(1).Row(i).Cell(5).Value).Trim().ToUpper().Split(' ')[2],
                                                Convert.ToString(book.Worksheet(1).Row(i).Cell(6).Value).Trim(),
                                                "",
                                                (svid) ? Convert.ToString(book.Worksheet(1).Row(i + 1).Cell(10).Value).Trim() : "",
                                                nzp_rod,
                                                Convert.ToString(book.Worksheet(1).Row(i).Cell(4).Value).Trim());
                        billKartDb.InsertGrgd(nzp_kart);
                    }

                }
            }
            book.Save();
        }
    }
}
