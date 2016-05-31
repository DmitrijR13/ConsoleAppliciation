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

namespace ConsoleApplication1.mainCode
{
    class BillUploadData
    {
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
}
