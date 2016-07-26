using ClosedXML.Excel;
using ConsoleApplication9;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1.mainCode
{
    class EzhkhReport
    {
        private pg pg;
        private Ora ora;
        public EzhkhReport()
        {
            pg = new pg();
            ora = new Ora();
        }

        public void RepPctHouse()
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
            wb.SaveAs(@"C:\temp\report" + DateTime.Now.ToShortDateString() + ".xlsx");
        }


        public void RepLift()
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

        public void RepCurRepair()
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

        public void ReportForMelnikov()
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

        public void ReportForBegin()
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

        public void AddHeadFio()
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

        public void AddKoap()
        {
            DataTable dt = ora.SelectKoap();
            var wb2 = new XLWorkbook(@"C:\Temp\2222.xlsx");

            for (int i = 8; i <= 1000; i++)//18953
            {
                if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim() == "Итого по району")
                    continue;
                if (i % 500 == 0)
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

        public void AddGkhCode2()
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

        public void AddGkhCode()
        {
            string[] stringSeparators = new string[] { "д." };
            int ws = 2;
            var wb1 = new XLWorkbook(@"C:\Temp\zhks(1).xlsx");
            for (int i = 2; i <= 2114; i++)//5
            {
                if (i % 50 == 0)
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
                        string dom = wb1.Worksheet(1).Row(i).Cell(1).Value.ToString().Trim().Replace("-", "");
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

        public void UpdateExel()
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

        public void GetChargeDataForEzhkh()
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

            wb.SaveAs(@"C:\temp\EzhkhImport_" + month + "_" + year + ".xlsx");
        }

        public void RepActcheck()
        {
            DataTable obj = pg.SelectActcheckInfo();
            var wb = new XLWorkbook(@"C:\temp\Копия Result2017OF.UL.IP.xlsx");
            string address = "";
            string municipality = "";
            int rowMove = 44;
            int colNum = 0;
            wb.Worksheet(1).Row(10).Cell(5).SetValue("161050630000565115");
            wb.Worksheet(1).Row(12).Cell(5).SetValue(obj.Rows[0][0] != null ? obj.Rows[0][0].ToString() : "");
            wb.Worksheet(1).Row(14).Cell(5).SetValue(@"12:00");
            wb.Worksheet(1).Row(16).Cell(5).Value = obj.Rows[0][1] != null ? obj.Rows[0][1].ToString() : "";
            wb.Worksheet(1).Row(18).Cell(5).SetValue(obj.Rows[0][0] != null ? obj.Rows[0][0].ToString() : "");
            wb.Worksheet(1).Row(20).Cell(5).SetValue(@"12:00");
            wb.Worksheet(1).Row(22).Cell(5).Value = "1";
            Int32 minutes1 = obj.Rows[0][3] != null && obj.Rows[0][3].ToString() != "" ? Convert.ToInt32(obj.Rows[0][3].ToString().Split(':')[0]) * 60 + Convert.ToInt32(obj.Rows[0][3].ToString().Split(':')[1]) : 0;
            Int32 minutes2 = obj.Rows[0][4] != null && obj.Rows[0][4].ToString() != "" ? Convert.ToInt32(obj.Rows[0][4].ToString().Split(':')[0]) * 60 + Convert.ToInt32(obj.Rows[0][4].ToString().Split(':')[1]) : 0;
            Decimal minutes3 = minutes2 - minutes1;
            String minutes4 = Math.Round(minutes3 / 60, 0) + "";
            wb.Worksheet(1).Row(24).Cell(5).Value = minutes3 > 0 ? minutes4 : "1";
            wb.Worksheet(1).Row(26).Cell(5).Value = obj.Rows[0][1] != null ? obj.Rows[0][1].ToString() : "";
            wb.Worksheet(1).Row(28).Cell(5).Value = obj.Rows[0][5] != null ? obj.Rows[0][5].ToString() : "";
            for (int i = 0; i < obj.Rows.Count; i++)
            {
                rowMove++;
                colNum++;
                wb.Worksheet(1).Row(rowMove).Cell(2).Value = colNum;
                wb.Worksheet(1).Row(rowMove).Cell(3).SetValue(obj.Rows[i][6] != null ? obj.Rows[i][6].ToString() : "");
                wb.Worksheet(1).Row(rowMove).Cell(4).SetValue(obj.Rows[i][7] != null ? obj.Rows[i][7].ToString() : "");
                wb.Worksheet(1).Row(rowMove).Cell(6).SetValue(obj.Rows[i][8] != null ? obj.Rows[i][8].ToString() : "");
                if(obj.Rows[i][8] != null && obj.Rows[i][8].ToString() != "")
                    wb.Worksheet(1).Row(rowMove).Cell(7).SetValue(obj.Rows[i][6] != null ? obj.Rows[i][6].ToString() : "");
                wb.Worksheet(1).Row(rowMove).Cell(8).SetValue(obj.Rows[i][9] != null ? obj.Rows[i][9].ToString() : "");
                wb.Worksheet(1).Row(rowMove).Cell(9).SetValue(obj.Rows[i][10] != null ? obj.Rows[i][10].ToString() : "");
                wb.Worksheet(1).Row(rowMove).Cell(10).SetValue(obj.Rows[i][11] != null ? obj.Rows[i][11].ToString() : "");
            }
            wb.SaveAs(@"C:\temp\161050630000565115.xls");
        }
    }
}
