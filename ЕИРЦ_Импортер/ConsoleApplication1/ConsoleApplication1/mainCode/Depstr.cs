using BytesRoad.Net.Ftp;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace ConsoleApplication1.mainCode
{
    class Depstr
    {
        private pg pg;
        public Depstr()
        {
            pg = new pg();
        }
        public void GetContractInfo()
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

        public void LoadDataFromFTPMore()
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
                for (int i = 0; i < 1; i++)
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

        public void LoadDataFromFTP()
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

        public void LoadDataFromFTP2()
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

        public void LoadDataFromFTP3()
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
    }
}
