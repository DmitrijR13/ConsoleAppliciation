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
    class BillPack
    {
        private pg pg;
        public BillPack()
        {
            pg = new pg();
        }

        public void ConvertFromAVB2()
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
                var paid = incomingData[2].Trim().Replace('.', ',');
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

        public void ConvertFromAVB()
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
        public void NccPack()
        {
            Int32 year = 2016;
            Int32 month = 2;
            BillPackDb billPackDb = new BillPackDb();
            var wb2 = new XLWorkbook(@"C:\temp\ЕИРЦ_Пачки\NCC_Out_num.xlsx");
            int num = Convert.ToInt32(wb2.Worksheet(1).Row(1).Cell(1).Value) + 1;
            wb2.Worksheet(1).Row(1).Cell(1).Value = num;
            wb2.Save();
            StreamWriter outPack = new StreamWriter(@"C:\Temp\ЕИРЦ_Пачки\ncc_out.txt", false, Encoding.GetEncoding("cp866"));
            outPack.WriteLine("<smpay_load_hdr><format_id>smpay_load_data</format_id><format_version>1</format_version><file_id>" + num + "</file_id></smpay_load_hdr>");
            List<string> prefs = new List<string>() { "bill01", "bill02" };
            DataTable dt = billPackDb.SelectSaldoForNCC(year, month, prefs);
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

        public void DymokPack()
        {
            BillPackDb billPackDb = new BillPackDb();
            Int32 year = 2016;
            Int32 month = 2;
            StreamWriter outPack = new StreamWriter(@"C:\Temp\dymok.csv", false, Encoding.UTF8);
            List<string> prefs = new List<string>() { "bill01", "bill02" };
            DataTable dt = billPackDb.SelectSaldoForAvtovazbank(year, month, prefs);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string eirc = Convert.ToString(dt.Rows[i][4]);
                string id = Convert.ToString(dt.Rows[i][0]);
                string address = Convert.ToString(dt.Rows[i][1]);
                string d = Math.Round(Convert.ToDecimal(dt.Rows[i][2]) > 0 ? Convert.ToDecimal(dt.Rows[i][2]) : 0, 2).ToString().Replace(',', '.');
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

        public void SberbankPack()
        {
            BillPackDb billPackDb = new BillPackDb();
            String fileName = "6321388192_40702810754400005587_001_y01";
            StreamWriter outPack = new StreamWriter(@"C:\Temp\" + fileName + ".txt", false, Encoding.GetEncoding("windows-1251"));
            string db = "192.168.1.25";
            List<string> prefs = new List<string>() { "bill01", "bill02" };
            DataTable dt = billPackDb.SelectSaldoForSberbank(db, 2, 2016, prefs);
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

        public void AvtovazbankPack()
        {
            BillPackDb billPackDb = new BillPackDb();
            Int32 year = 2016;
            Int32 month = 2;
            StreamWriter outPack = new StreamWriter(@"C:\Temp\avtovazbank.csv", false, Encoding.UTF8);
            List<string> prefs = new List<string>() { "bill01", "bill02" };
            DataTable dt = billPackDb.SelectSaldoForAvtovazbank2(year, month, prefs);
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
    }
}
