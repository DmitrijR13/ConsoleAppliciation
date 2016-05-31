using BytesRoad.Net.Ftp;
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
