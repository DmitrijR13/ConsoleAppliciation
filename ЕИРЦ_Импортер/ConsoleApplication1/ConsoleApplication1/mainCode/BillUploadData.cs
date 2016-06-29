using ClosedXML.Excel;
using ConsoleApplication1.Database;
using ConsoleApplication9;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ConsoleApplication1.mainCode
{
    class BillUploadData
    {
        private pg pg;
        public BillUploadData()
        {
            pg = new pg();
        }

        public void AddPkod()
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

        public void CreateXmlByCountersVal()
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

        public void CompareDate()
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

        public void CreateScript()
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
                              "-01'::date AND dat_oper < '2016-" + (i + 1).ToString("00") +
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

        public void CorrectPeni()
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

        public void RepChargePrihByMonth()
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
                                            delegate (Service s)
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
                                            delegate (Service s)
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
                                            delegate (Service s)
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
                                            delegate (Service s)
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

        public void GetSaldoWithParam()
        {
            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Лист1");
            var tableSum = pg.GetSaldoAndParam();
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

        public void GetCountersTlt()
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

        public void CreateKvarFile()
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
            BillUploadDataDb billUploadDataDb = new BillUploadDataDb();
            for (int j = 0; j < dt1.Rows.Count; j++)
            {
                dtHouse = billUploadDataDb.SelectHouseCode(dt1.Rows[j][0].ToString());


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
                dtPeople = billUploadDataDb.SelectLN4Code(dt1.Rows[j][0].ToString());
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
