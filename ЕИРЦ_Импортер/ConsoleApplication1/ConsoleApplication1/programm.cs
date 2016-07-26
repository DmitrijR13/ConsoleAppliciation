using BytesRoad.Net.Ftp;
using ClosedXML.Excel;
using ConsoleApplication1.Database;
using ConsoleApplication1.mainCode;
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
using System.Data.OracleClient;
using System.Security.Cryptography.X509Certificates;

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
            BillBaseDb billBaseDb = new BillBaseDb();
            EzhkhBaseDb ezhkhBaseDb = new EzhkhBaseDb();
            InsertPeople ipProg = new InsertPeople();
            InsertPeopleDb insPeopleDb = new InsertPeopleDb();
            Depstr depstr = new Depstr();
            int type;
            Console.WriteLine("17 = Акты из ЖКХ в шаблон Эксель");
            Console.Write("Введите тип операции:");
            type = Convert.ToInt32(Console.ReadLine());

            #region 5 Update MKD Area
            if (type == 5)
            {
                EzhkhInsertData data = new EzhkhInsertData();
                data.UpdateAreaMkd();
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

            #region 7 Free
            else if (type == 7)
            {
                WebRequest request = WebRequest.Create("http://85.140.61.250/GkhService/Service1.svc/GetDataCSV?token=f8f84d10cc6727b20becb7c5e85de047");
                request.Method = "GET";
                WebResponse response = request.GetResponse();

            }
            #endregion

            #region 8 Free
            else if (type == 8)
            {
                Dictionary<string, Int32> month = new Dictionary<string, Int32>();
                month.Add("январь", 1);
                month.Add("февраль", 2);
                month.Add("март", 3);
                month.Add("апрель", 4);
                month.Add("май", 5);
                month.Add("июнь", 6);
                month.Add("июль", 7);
                month.Add("август", 8);
                month.Add("сентябрь", 9);
                month.Add("октябрь", 10);
                month.Add("ноябрь", 11);
                month.Add("декабрь", 12);
                List<Int32> clearHouse = new List<Int32>();
                var wb2 = new XLWorkbook(@"C:\temp\ЭЖКХ 1 квартал ЖКС 2016.xlsx");
                for (int i = 3; i <= 6; i++)
                {
                    if (Convert.ToString(wb2.Worksheet(2).Row(i).Cell(1).Value) != "")
                    {
                        Int32 roId = ezhkhBaseDb.SelectRoIDByGkhCode(Convert.ToString(wb2.Worksheet(2).Row(i).Cell(1).Value));
                        if (roId == 0 || roId == 2)
                        {
                            wb2.Worksheet(2).Row(i).Cell(11).Value = roId;
                            wb2.Worksheet(2).Row(i).Style.Fill.BackgroundColor = XLColor.Yellow;
                            continue;
                        }                         
                        if (!clearHouse.Contains(roId))
                        {
                            clearHouse.Add(roId);
                            ezhkhBaseDb.DelCurRepair(roId);
                        }
                        List<string> repId = ezhkhBaseDb.SelectCurRepWorkId(Convert.ToString(wb2.Worksheet(2).Row(i).Cell(2).Value));
                        if(repId == null)
                        {
                            wb2.Worksheet(2).Row(i).Style.Fill.BackgroundColor = XLColor.Red;
                            continue;
                        }
                        String planDate = month.ContainsKey(Convert.ToString(wb2.Worksheet(2).Row(i).Cell(3).Value))
                            ? "2016-" + month[Convert.ToString(wb2.Worksheet(2).Row(i).Cell(3).Value)] + "-01" : "";
                        String planWork = Convert.ToString(wb2.Worksheet(2).Row(i).Cell(4).Value) == "" ? "0" : Convert.ToString(wb2.Worksheet(2).Row(i).Cell(4).Value).Replace(",",".");
                        String planSum = Convert.ToString(wb2.Worksheet(2).Row(i).Cell(6).Value) == "" ? "0" : Convert.ToString(wb2.Worksheet(2).Row(i).Cell(6).Value).Replace(",", ".");

                        String factDate = month.ContainsKey(Convert.ToString(wb2.Worksheet(2).Row(i).Cell(7).Value))
                            ? "2016-" + month[Convert.ToString(wb2.Worksheet(2).Row(i).Cell(7).Value)] + "-01" : "";
                        String factWork = Convert.ToString(wb2.Worksheet(2).Row(i).Cell(8).Value) == "" ? "0" : Convert.ToString(wb2.Worksheet(2).Row(i).Cell(8).Value).Replace(",", ".");
                        String factSum = Convert.ToString(wb2.Worksheet(2).Row(i).Cell(10).Value) == "" ? "0" : Convert.ToString(wb2.Worksheet(2).Row(i).Cell(10).Value).Replace(",", ".");

                        ezhkhBaseDb.InsertCurRepair(roId, factDate, factSum, factWork, planDate, planSum, planWork, repId[1], repId[0]);
                    }
                }
            }
            #endregion

            #region 9 Проверка BarCode
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

            #region 13 Insert Blob
            else if (type == 13)
            {
                //Step 1
                // Connect to database
                // Note: Modify User Id, Password, Data Source as per your database setup
                string constr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=85.140.61.250)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ezhkh)));User Id = blob; Password = ACTANONVERBA";

                OracleConnection con = new OracleConnection(constr);
                con.Open();
                Console.WriteLine("Connected to database!");

                // Step 2
                // Note: Modify the Source and Destination location
                // of the image as per your machine settings
                String SourceLoc = @"C:\Temp\Cer\certificate.cer";

                // provide read access to the file

                FileStream fs = new FileStream(SourceLoc, FileMode.Open, FileAccess.Read);

                // Create a byte array of file stream length
                byte[] ImageData = new byte[fs.Length];

                //Read block of bytes from stream into the byte array
                fs.Read(ImageData, 0, System.Convert.ToInt32(fs.Length));

                //Close the File Stream
                fs.Close();

                // Step 3
                // Create Anonymous PL/SQL block string
                String block = " BEGIN " +
                               " INSERT INTO test (test) VALUES (:1); " +
                               //" SELECT test into :2 from test; " +
                               " END; ";

                // Set command to create Anonymous PL/SQL Block
                OracleCommand cmd = new OracleCommand();
                cmd.CommandText = block;
                cmd.Connection = con;


                // Since executing an anonymous PL/SQL block, setting the command type
                // as Text instead of StoredProcedure
                cmd.CommandType = CommandType.Text;

                // Bind the parameter as OracleDbType.Blob to command for retrieving the image
                cmd.Parameters.AddWithValue("1", ImageData);

                // Step 5
                // Execute the Anonymous PL/SQL Block

                // The anonymous PL/SQL block inserts the image to the
                // database and then retrieves the images as an output parameter
                cmd.ExecuteNonQuery();
                Console.WriteLine("Image file inserted to database from " + SourceLoc);

                
                con.Close();
            }
            #endregion

            #region 14 Free
            else if (type == 14)
            {
                byte[] file;
                using (var stream = new FileStream(@"C:\Temp\Cer\algoritm.pdf", FileMode.Open, FileAccess.Read))
                {
                    using (var reader = new BinaryReader(stream))
                    {
                        file = reader.ReadBytes((int)stream.Length);
                    }
                }
                string constr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=85.140.61.250)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ezhkh)));User Id = blob; Password = ACTANONVERBA";
                OracleConnection conn = new OracleConnection(constr);
                string cmdText1 = "INSERT INTO test (id, test) VALUES (2, :1)";
                OracleCommand cmd1 = new OracleCommand(cmdText1, conn);
                cmd1.Parameters.AddWithValue("1", file);
                conn.Open();
                cmd1.ExecuteNonQuery();
                conn.Close();
            }
            #endregion

            #region 11 Free
            else if (type == 11)
            {
                string connectionString; string query; string cert;

                connectionString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=85.140.61.250)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ezhkh)));User Id = blob; Password = ACTANONVERBA;";
                query = "SELECT test FROM test";

                using (OracleConnection cn = new OracleConnection(connectionString))
                {
                    OracleCommand cmd = new OracleCommand(query, cn);
                    cn.Open();
                    byte[] b = (byte[])cmd.ExecuteScalar();
                    cert = System.Text.Encoding.ASCII.GetString(b);
                }
                int bufferSize = 100000000;                   // Size of the BLOB buffer.
                byte[] outbyte = new byte[bufferSize];
                X509Certificate2 serverCert = new X509Certificate2(Encoding.ASCII.GetBytes(cert));
                outbyte = serverCert.Export(X509ContentType.Cert);     
                File.WriteAllBytes(@"C:\Temp\Cer\6.cer", outbyte); // Requires System.IO
            }
            #endregion

            #region 15 Free
            else if (type == 15)
            {
                string connectionString; string query; string cert;
                byte[] b;

                connectionString = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=85.140.61.250)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ezhkh)));User Id = blob; Password = ACTANONVERBA;";
                query = "SELECT test FROM test where id = 2";

                using (OracleConnection cn = new OracleConnection(connectionString))
                {
                    OracleCommand cmd = new OracleCommand(query, cn);
                    cn.Open();
                    b = (byte[])cmd.ExecuteScalar();
                }
                File.WriteAllBytes(@"C:\Temp\Cer\algor.pdf", b); // Requires System.IO
            }
            #endregion

            #region 16 Free
            else if (type == 16)
            {
                EzhkhInsertDataDb ezhkhInsertDataDb = new EzhkhInsertDataDb();
                var book = new XLWorkbook(@"C:\temp\Копия Капитальный ремонт мкд.xlsx");
                for (int i = 1; i <= 373; i++)
                {
                    string str = ezhkhInsertDataDb.UpdatePhysicalWearTehPassport(Convert.ToString(book.Worksheet(1).Row(i).Cell(1).Value).Trim(),
                        Convert.ToDecimal(book.Worksheet(1).Row(i).Cell(2).Value));
                    if (str.Split('|')[0] != "ЗАГРУЖЕНО")
                        Console.WriteLine(str);
                    else
                    {
                        if(str.Split('|')[1] == "0")
                            Console.WriteLine(Convert.ToString(book.Worksheet(1).Row(i).Cell(1).Value).Trim() + " = " + str.Split('|')[1]);
                    }
                }
                book.Save();
            }
            #endregion

            #region 17 Акты из ЖКХ в шаблон Эксель
            else if (type == 17)
            {
                EzhkhReport ezhkhReport = new EzhkhReport();
                ezhkhReport.RepActcheck();
            }
            #endregion

            #region 20 Free
            else if (type == 20)
            {
               
            }
            #endregion

            #region 18 Free
            else if (type == 18)
            {
                
            }
            #endregion

            #region 19 Free
            else if (type == 19)
            {
                
            }
            #endregion

            #region 10 Free
            else if (type == 10)
            {
                
            }
            #endregion

            #region 21 Free
            else if (type == 21)
            {
               
            }
            #endregion

            #region 22 Free
            else if (type == 22)
            {

            }
            #endregion

            #region 23 Free
            else if (type == 23)
            {
               
            }
            #endregion

            #region 24 Free
            else if (type == 24)
            {
               
            }
            #endregion

            #region 25 Free
            else if (type == 25)
            {
                
            }
            #endregion

            #region 26 Free
            else if (type == 26)
            {
                
            }
            #endregion

            #region 4 Free
            else if (type == 4)
            {
                
            }
            #endregion

            #region 3 Free
            else if (type == 3)
            {
               
            }
            #endregion

            #region 2 Free
            else if (type == 2)
            {
                
            }
            #endregion

            #region 1 Free
            if (type == 1)//запись из Эксельки
            {
                
            }
            #endregion

            #region 0 формирование текстового файла для загрузки в биллинг
            else if (type == 0)
            {
                BillUploadData billUploadData = new BillUploadData();
                billUploadData.CreateKvarFile();
            }
            #endregion

            #region 27 Free
            else if (type == 27)
            {
                
            }
            #endregion

            #region 30 InsertHouseManOrg
            else if (type == 30)
            {
                EzhkhInsertData ezhkhInsertData = new EzhkhInsertData();
                ezhkhInsertData.AddHouseManOrg();
            }
            #endregion

            #region 310 InsertHouseManOrg
            else if (type == 310)
            {
                EzhkhInsertData ezhkhInsertData = new EzhkhInsertData();
                ezhkhInsertData.AddHouseManOrgFromFile();
            }
            #endregion

            //Поставщики коммунальных услуг
            #region 31 InsertCommunalOrg
            else if (type == 31)
            {
                EzhkhInsertData ezhkhInsertData = new EzhkhInsertData();
                ezhkhInsertData.AddCommunalOrg();
            }
            #endregion

            #region 42 InsertResOrg
            else if (type == 42)
            {
                EzhkhInsertData ezhkhInsertData = new EzhkhInsertData();
                ezhkhInsertData.AddResOrg();
            }
            #endregion

            #region 32
            else if (type == 32)
            {
                InsertPeople insertPeople = new InsertPeople();
                insertPeople.Ins1();
            }
            #endregion

            #region 33
            else if (type == 33)
            {
                InsertPeople insertPeople = new InsertPeople();
                insertPeople.Ins2();
            }
            #endregion

            #region 34
            else if (type == 34)
            {
                InsertPeople insertPeople = new InsertPeople();
                insertPeople.Ins3();
            }
            #endregion

            #region 35
            else if (type == 35)
            {
                InsertPeople insertPeople = new InsertPeople();
                insertPeople.Ins4();
            }
            #endregion

            #region 36 Free
            else if (type == 36)
            {
                
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

            #region 38 Free
            else if (type == 38)
            {
                
            }
            #endregion

            #region 39 Free
            else if (type == 39)
            {
               
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
                EzhkhReport ezhkhReport = new EzhkhReport();
                ezhkhReport.RepCurRepair();
            }
            #endregion

            #region 43 Отчет по лифтам
            else if (type == 43)
            {
                EzhkhReport ezhkhReport = new EzhkhReport();
                ezhkhReport.RepLift();
            }
            #endregion

            #region 44 Отчет по проценту заполнения домов
            else if (type == 44)
            {
                EzhkhReport ezhkhReport = new EzhkhReport();
                ezhkhReport.RepPctHouse();
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

            #region 49 Free
            else if (type == 49)
            {
                
            }
            #endregion

            #region 50 Free
            else if (type == 50)
            {
               
            }
            #endregion

            //Перенос тарифов из Информикса в Постгре
            #region 51
            else if (type == 51)
            {
                BillInsertData billInsertData = new BillInsertData();
                billInsertData.InsertTarif();
            }
            #endregion

            //Перенос нормативов из Информикса в Постгре
            #region 52
            else if (type == 52)
            {
                BillInsertData billInsertData = new BillInsertData();
                billInsertData.InsertNormativ();
            }
            #endregion

            //Перенос нормативов из Информикса в Постгре
            #region 53
            else if (type == 53)
            {
                BillInsertData billInsertData = new BillInsertData();
                billInsertData.InsertNormativ2();
            }
            #endregion

            #region 54 Free
            else if (type == 54)
            {
               
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
                Console.Write("Введите наименование БД:");
                string database = Console.ReadLine();
                var wb2 = new XLWorkbook(@"C:\temp\Копия Первомайская 7-загрузить.xlsx");
                string[] stringSeparators = new string[] { "в." };
                string[] stringSeparatorsUl = new string[] { ", ул." };
                string[] stringSeparatorsDom = new string[] { ", д." };
                string address = "";
                int nzp_ul = 0;
                string dom = "";
                int nzp_dom = 0;
                Dictionary<string, string> addHouses = new Dictionary<string, string>();
                for (int i = 2; i <= 20; i++)
                {
                    if (address != Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim())
                    {
                        address = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim();

                        string rajon = pg.SelectNzpRaj(database, address.Trim().Split(stringSeparatorsUl, StringSplitOptions.None)[0].ToUpper());
                        if (rajon.Split('|')[0] == "0")
                        {
                            Console.WriteLine(address + ": " + rajon.Split('|')[1]);
                            nzp_dom = 0;
                        }
                        else
                        {
                            string ulica = pg.SelectNzpUl(database, 
                                address.Trim().Split(stringSeparatorsUl, StringSplitOptions.None)[1].Split(stringSeparatorsDom, StringSplitOptions.None)[0].ToUpper().Trim(), 
                                rajon.Split('|')[0]);
                            if (ulica.Split('|')[0] == "0")
                            {
                                Console.WriteLine(address + ": " + ulica.Split('|')[1]);
                                nzp_dom = 0;
                            }
                            else
                            {
                                nzp_ul = Convert.ToInt32(ulica.Split('|')[0]);
                                Console.WriteLine(address + ": " + nzp_ul);
                                dom = Convert.ToString(address.Trim().Split(stringSeparatorsUl, StringSplitOptions.None)[1].Split(stringSeparatorsDom, StringSplitOptions.None)[1]).Trim().ToUpper();
                                nzp_dom = pg.InsertDom(database, nzp_ul, dom);

                                Console.WriteLine(address + ": " + nzp_dom);
                                for (int j = 2; j <= 37; j++)
                                {
                                    if(Convert.ToString(wb2.Worksheet(2).Row(j).Cell(1).Value).Trim() == address)
                                    {
                                        if(Convert.ToString(wb2.Worksheet(2).Row(j).Cell(2).Value).Trim() != "")
                                            pg.InsertDomPrm(database, nzp_dom, Convert.ToString(wb2.Worksheet(2).Row(j).Cell(2).Value).Trim(), 40);
                                        if (Convert.ToString(wb2.Worksheet(2).Row(j).Cell(3).Value).Trim() != "")
                                            pg.InsertDomPrm(database, nzp_dom, Convert.ToString(wb2.Worksheet(2).Row(j).Cell(3).Value).Trim(), 150);
                                        if (Convert.ToString(wb2.Worksheet(2).Row(j).Cell(4).Value).Trim() != "")
                                            pg.InsertDomPrm(database, nzp_dom, Convert.ToString(wb2.Worksheet(2).Row(j).Cell(4).Value).Trim(), 37);
                                        if (Convert.ToString(wb2.Worksheet(2).Row(j).Cell(5).Value).Trim() != "")
                                            pg.InsertDomPrm(database, nzp_dom, Convert.ToString(wb2.Worksheet(2).Row(j).Cell(5).Value).Trim(), 2049);
                                        break;
                                    }
                                        
                                }
                            }
                        }
                        
                    }
                    /*if (dom != Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim())
                    {
                        dom = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim();
                        nzp_dom = pg.InsertDom(nzp_ul, dom);
                        Console.WriteLine(dom + ": " + nzp_dom);
                    }*/
                    string nkvar = "";
                    //if(Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim().Substring(0,3) == "Кв.")
                    nkvar = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim();
                    //else
                     //   nkvar =Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim();
                    //int num_ls = 0;
                    //int number;
                    //int subLs = 0;
                    //string num_ls_full = Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim().Split('№')[1].Trim();
                    //for (int j = 0; j < num_ls_full.Length; j++)
                    //{
                    //    if (Int32.TryParse(num_ls_full.Substring(j, 1), out number))
                    //        subLs++;
                    //    else
                    //        break;
                    //}
                    //if(subLs > 0)
                    //    num_ls = Convert.ToInt32(num_ls_full.Substring(0, subLs));

                    //int ikvar = 0;
                    //int subKvar = 0;
                    //for (int j = 0; j < nkvar.Length; j++)
                    //{
                    //    if (Int32.TryParse(nkvar.Substring(j, 1), out number))
                    //        subKvar++;
                    //    else
                    //        break;
                    //}
                    //if (subKvar > 0)
                    //    ikvar = Convert.ToInt32(nkvar.Substring(0, subKvar));
                    if(nzp_dom != 0)
                    {
                        int nzp_kvar = pg.InsertKvar(database, nzp_dom.ToString(), Convert.ToString(wb2.Worksheet(1).Row(i).Cell(7).Value).Trim(), nkvar);
                        Console.WriteLine(address + ", кв. " + nkvar + ": " + nzp_kvar);
                        pg.InsertDateOpen(database, nzp_kvar, "01.07.2016");
                        if (wb2.Worksheet(1).Row(i).Cell(3).Value != null && Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim() != "")
                            pg.InsertPrm1(database, nzp_kvar, Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim(), 4);
                        if (wb2.Worksheet(1).Row(i).Cell(4).Value != null && Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim() != "")
                            pg.InsertPrm1(database, nzp_kvar, Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim(), 6);
                        if (wb2.Worksheet(1).Row(i).Cell(5).Value != null && Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value).Trim() != "")
                            pg.InsertPrm1(database, nzp_kvar, Convert.ToString(wb2.Worksheet(1).Row(i).Cell(5).Value).Trim(), 5);
                        if (wb2.Worksheet(1).Row(i).Cell(6).Value != null && Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value).Trim() != "")
                        {
                            if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value).Trim() == "Да")
                                pg.InsertPrm1(database, nzp_kvar, "1", 8);
                            else if (Convert.ToString(wb2.Worksheet(1).Row(i).Cell(6).Value).Trim() == "Нет")
                                pg.InsertPrm1(database, nzp_kvar, "0", 8);
                        }
                    }     
                }
            }
            #endregion

            //Перенос домов и жильцов из Информикса в Постгри
            #region 59 Free
            else if (type == 59)
            {
               
            }
            #endregion

            //Групповой ввод характеристик жилья
            #region 60
            else if (type == 60)
            {
                BillInsertData billInsertData = new BillInsertData();
                billInsertData.GroupHouseConstruct();
            }
            #endregion

            #region 61 Free
            else if (type == 61)
            {
                
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

            //загрузка ПС КАРТОТЕКА
            #region 65
            else if (type == 65)
            {
                BillKart billKart = new BillKart();
                billKart.LoadKart();
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
            #region 74 стройка. проверка контрактов
            else if (type == 74)
            {
                depstr.GetContractInfo();
            }
            #endregion

            #region 75 Free
            else if (type == 75)
            {
                
            }
            #endregion

            #region 76 Free
            else if (type == 76)
            {
                
            }
            #endregion

            #region 77 Free
            else if (type == 77)
            {

            }
            #endregion

            #region 78 Мельникову
            else if (type == 78)
            {
                EzhkhReport ezhkhReport = new EzhkhReport();
                ezhkhReport.ReportForMelnikov();
            }
            #endregion

            //ftp connect
            #region 79
            else if (type == 79)
            {
                depstr.LoadDataFromFTP3();
            }
            #endregion

            //Стройка. ftp connect
            #region 80
            else if (type == 80)
            {
                depstr.LoadDataFromFTP2();        
            }
            #endregion

            //Стройка. ftp connect
            #region 83
            else if (type == 83)
            {
                depstr.LoadDataFromFTPMore();
            }
            #endregion

            #region 830 Free
            else if (type == 830)
            {
                
            }
            #endregion

            #region 82 Free
            else if (type == 82)
            {

            }
            #endregion

            //Стройка. Загрузка с FTP
            #region 91
            else if (type == 91)
            {
                depstr.LoadDataFromFTP();
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
                EzhkhReport ezhkhReport = new EzhkhReport();
                ezhkhReport.ReportForBegin();
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
                BillUploadData billUploadData = new BillUploadData();
                billUploadData.AddPkod();
            }
            #endregion

            //обновление даты оплат
            #region 95
            else if (type == 95)
            {
                BillInsertData billInsertData = new BillInsertData();
                billInsertData.UpdatePackDate();
            }
            #endregion

            //Собственники с пробегом по всем файлам
            #region 96
            else if (type == 96)
            {
                DirectoryInfo dir = new DirectoryInfo(@"C:\temp\houses");
                string[] stringSeparators = new string[] { "Кв." };
                foreach (var item in dir.GetFiles())
                {
                    Console.WriteLine(item.Name);
                    var wb2 = new XLWorkbook(@"C:\temp\houses\" + item.Name);
                    for (int i = 3; i <= 1000; i++)
                    {
                        if (wb2.Worksheet(1).Row(i).Cell(1).Value == null || Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim() == "")
                            break;
                        try
                        {
                            string str = insPeopleDb.InsertPeople5(item.Name.Split('.')[0],
                                Convert.ToString(wb2.Worksheet(1).Row(i).Cell(2).Value).Trim(),
                                Convert.ToString(wb2.Worksheet(1).Row(i).Cell(3).Value).Trim(),
                                "",
                                Convert.ToString(wb2.Worksheet(1).Row(i).Cell(4).Value).Trim(),
                                "",
                                Convert.ToString(wb2.Worksheet(1).Row(i).Cell(1).Value).Trim());
                            if (str != "ЗАГРУЖЕНО")
                                Console.WriteLine(wb2.Worksheet(1).Row(i).Cell(1).Value + ":" + str);
                        }
                        catch(Exception e)
                        {
                            Console.WriteLine(wb2.Worksheet(1).Row(i).Cell(1).Value + ":" + e.ToString());
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
                    string kvar = billBaseDb.SelectNzpKvarByKvarDom(database, wb2.Worksheet(1).Row(i).Cell(1).Value.ToString(), 7155108);
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

            #region 104 Фамилии по должностям
            else if (type == 104)
            {
                EzhkhReport ezhkhReport = new EzhkhReport();
                ezhkhReport.AddHeadFio();
            }
            #endregion

            //Обновление площадей по ЛС
            #region 105 UpdateKvarTotalSquare
            else if (type == 105)
            {
                BillInsertData billInsertData = new BillInsertData();
                billInsertData.UpdateKvarTotalSquare();
            }
            #endregion

            //Загрузка счетчиком для сайта
            #region 106 Формирование xml для выгрузки счетчиков на сайт
            else if (type == 106)
            {
                BillUploadData billUploadData = new BillUploadData();
                billUploadData.CreateXmlByCountersVal();
            }
            #endregion

            //Импорт квартир в РЦ
            #region 107
            else if (type == 107)
            {
                BillInsertData billInsertData = new BillInsertData();
                billInsertData.ImportKvarFromExcel();
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

            //Добавляем показания счетчиков
            #region 109
            else if (type == 109)
            {
                BillInsertData billInsertData = new BillInsertData();
                billInsertData.InsertCountersVal3();
            }
            #endregion

            #region 1091 Free
            else if (type == 1091)
            {

            }
            #endregion

            #region 110 Free
            else if (type == 110)
            {
                
            }
            #endregion

            //Update даты закрытия счетчиков
            #region 111
            else if (type == 111)
            {
                BillInsertData billInsertData = new BillInsertData();
                billInsertData.UpdateCountersDatClose();
            }
            #endregion

            //Добавление показаний счетчиков
            #region 112
            else if (type == 112)
            {
                BillInsertData billInsertData = new BillInsertData();
                billInsertData.InsertCountersVal2();
            }
            #endregion

            //Проверка(добавление) показания счетчиков
            #region 113
            else if (type == 113)
            {
                BillInsertData billInsertData = new BillInsertData();
                billInsertData.CheckOrInsertCountersVal();
            }
            #endregion

            //Добавляем показания счетчиков
            #region 1130
            else if (type == 1130)
            {
                BillInsertData billInsertData = new BillInsertData();
                billInsertData.InsertCountersVal();
            }
            #endregion

            #region 114 Free
            else if (type == 114)
            {
                
            }
            #endregion

            //Сальдо по пени
            #region 115
            else if (type == 115)
            {
                BillInsertData billInsertData = new BillInsertData();
                billInsertData.UpdateOutSaldo();
            }
            #endregion

            //Загрузка ПС Картотека
            #region 116
            else if (type == 116)
            {
                BillKart billKart = new BillKart();
                billKart.LoadKart2();
            }
            #endregion

            #region ПАЧКИ ОПЛАТ ЕИРЦ
            #region 1000 Формирование оплат для NCC
            else if (type == 1000)
            {
                BillPack bp = new BillPack();
                bp.NccPack();
            }
            #endregion

            #region 1100 Формирование оплат для Дымка
            else if (type == 1100)
            {
                BillPack bp = new BillPack();
                bp.DymokPack();
            }
            #endregion

            #region 1200 Формирование оплат для Сбербанка
            else if (type == 1200)
            {
                BillPack bp = new BillPack();
                bp.SberbankPack();
            }
            #endregion

            #region 1300 Формирование оплат для Автовазбанк
            else if (type == 1300)
            {
                BillPack bp = new BillPack();
                bp.AvtovazbankPack();
            }
            #endregion
            #endregion

            #region 1110 пачки из АВБ
            else if (type == 1110)
            {
                BillPack billPack = new BillPack();
                billPack.ConvertFromAVB2();
            }
            #endregion

            #region 1111 пачки из АВБ
            else if (type == 1111)
            {
                BillPack billPack = new BillPack();
                billPack.ConvertFromAVB();
            }
            #endregion

            #region 117 проставляем статьи КОАП
            else if (type == 117)
            {
                EzhkhReport ezhkhReport = new EzhkhReport();
                ezhkhReport.AddKoap();
            }
            #endregion

            #region 118 сравнение показаний в эксельках
            else if (type == 118)
            {
                BillUploadData billUploadData = new BillUploadData();
                billUploadData.CreateScript();
            }
            #endregion

            #region 119 проставляем GKH_CODE
            else if (type == 119)
            {
                EzhkhReport ezhkhReport = new EzhkhReport();
                ezhkhReport.AddGkhCode2();
            }
            #endregion

            #region 120 Сумма по двум эксель файлам
            else if (type == 120)
            {
                ExcelOperation exelOperation = new ExcelOperation();
                exelOperation.Sum2File();              
            }
            #endregion

            #region 121 Делим значения на 2 в Эксель
            else if (type == 121)
            {
                ExcelOperation exelOperation = new ExcelOperation();
                exelOperation.DevideByTwo();
            }
            #endregion

            #region 122 Формируем скрипт из txt
            else if (type == 122)
            {
                BillUploadData billUploadData = new BillUploadData();
                billUploadData.CreateScript();
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
                EzhkhReport ezhkhReport = new EzhkhReport();
                ezhkhReport.GetChargeDataForEzhkh();
            }
            #endregion

            #region 126 фаил для Бегина
            else if (type == 126)
            {
                EzhkhReport ezhkhReport = new EzhkhReport();
                ezhkhReport.UpdateExel();
            }
            #endregion

            #region 127 Тест Convert.ToDecimal()
            else if (type == 127)
            {
                Decimal d1 = Convert.ToDecimal("0,00");
                Console.WriteLine(d1);
            }
            #endregion

            #region 128 фаил для Бегина
            else if (type == 128)
            {
                EzhkhReport ezhkhReport = new EzhkhReport();
                ezhkhReport.AddGkhCode();
            }
            #endregion

            #region 129 корректировка по пеням
            else if (type == 129)
            {
                BillUploadData billUploadData = new BillUploadData();
                billUploadData.CorrectPeni();
            }
            #endregion

            #region 130 Складываем отчет по начислению и оплате за месяца
            else if (type == 130)
            {
                BillUploadData billUploadData = new BillUploadData();
                billUploadData.RepChargePrihByMonth();
            }
            #endregion

            #region 131 Тлт выгрузка
            else if (type == 131)
            {
                BillUploadData billUploadData = new BillUploadData();
                billUploadData.GetSaldoWithParam();
            }
            #endregion

            #region 132 Тлт выгрузка счетчики
            else if (type == 132)
            {
                BillUploadData billUploadData = new BillUploadData();
                billUploadData.GetCountersTlt();
            }
            #endregion

            #region 133 Free
            else if (type == 133)
            {
                
            }
            #endregion

            #region 134 Free
            else if (type == 134)
            {
                
            }
            #endregion

            #region 135 Проставка и замена ezhkh_code
            else if (type == 135)
            {
                EzhkhInsertData ezhkhInsertData = new EzhkhInsertData();
                ezhkhInsertData.UpdateGkhCode();
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

        
    }   
}