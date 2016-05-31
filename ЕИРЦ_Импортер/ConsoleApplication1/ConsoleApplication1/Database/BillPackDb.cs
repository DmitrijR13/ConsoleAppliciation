using Npgsql;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1.Database
{
    class BillPackDb
    {
        public DataTable SelectSaldoForNCC(Int32 year, Int32 month, List<string> prefs)
        {
            string connStr = "Server=192.168.1.25;Database=billTlt;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("pkod");
            dataTable.Columns.Add("qwerty");
            dataTable.Columns.Add("sum_outsaldo");
            DataRow row2;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            foreach (string pref in prefs)
            {
                string cmdText = @"SELECT k.pkod, ul.ulicareg || ' ' || ul.ulica || ' д.' || d.ndom || ' кв ' || k.ikvar || ' к ' || CASE WHEN k.nkvar_n = '' 
                                OR k.nkvar_n  is null THEN '-' ELSE k.nkvar_n END as qwerty, c.sum_outsaldo 
                                FROM " + pref + @"_charge_" + (year - 2000).ToString("00") + @".charge_" +
                                       (pref == "bill02" ? (month - 1).ToString("00") : month.ToString("00")) +
                                @" c 
                                INNER JOIN " + pref + @"_data.kvar k on k.nzp_kvar = c.nzp_kvar 
                                INNER JOIN " + pref + @"_data.dom d on d.nzp_dom = k.nzp_dom
                                INNER JOIN " + pref + @"_data.s_ulica ul on ul.nzp_ul = d.nzp_ul
                                where nzp_serv = 1";
                NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
                NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
                DataTable dt = new DataTable();

                try
                {
                    da.Fill(dt);
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        row2 = dataTable.NewRow();
                        row2["pkod"] = dt.Rows[i][0];
                        row2["qwerty"] = dt.Rows[i][1];
                        row2["sum_outsaldo"] = dt.Rows[i][2];
                        dataTable.Rows.Add(row2);
                    }
                }
                catch (Exception)
                {

                }
            }
            return dataTable;
        }

        public DataTable SelectSaldoForAvtovazbank(Int32 year, Int32 month, List<string> prefs)
        {
            string connStr = "Server=192.168.1.25;Database=billTlt;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("pkod");
            dataTable.Columns.Add("qwerty");
            dataTable.Columns.Add("sum_outsaldo");
            dataTable.Columns.Add("fio");
            dataTable.Columns.Add("rekvizit");
            DataRow row2;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            foreach (string pref in prefs)
            {
                string rekvizit = "";
                switch (pref)
                {
                    case "bill01":
                        {
                            rekvizit = "ТСЖ-1;6321195867;";
                            break;
                        }
                    case "bill02":
                        {
                            rekvizit = "ТЭВИС;6320000561;";
                            break;
                        }
                }
                string cmdText = @"SELECT k.pkod, ul.ulicareg || ' ' || ul.ulica || ' д.' || d.ndom || ' кв ' || k.ikvar || ' к ' || CASE WHEN k.nkvar_n = '' 
                                OR k.nkvar_n  is null THEN '-' ELSE k.nkvar_n END as qwerty, sum_outsaldo, k.fio 
                                FROM " + pref + @"_charge_" + (year - 2000).ToString("00") + @".charge_" +
                                       (pref == "bill02" ? (month - 1).ToString("00") : month.ToString("00")) +
                                 @" c 
                                INNER JOIN " + pref + @"_data.kvar k on k.nzp_kvar = c.nzp_kvar 
                                INNER JOIN " + pref + @"_data.dom d on d.nzp_dom = k.nzp_dom
                                INNER JOIN " + pref + @"_data.s_ulica ul on ul.nzp_ul = d.nzp_ul
                                where nzp_serv = 1";
                NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
                NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
                DataTable dt = new DataTable();

                try
                {
                    da.Fill(dt);
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        row2 = dataTable.NewRow();
                        row2["pkod"] = dt.Rows[i][0];
                        row2["qwerty"] = dt.Rows[i][1];
                        row2["sum_outsaldo"] = dt.Rows[i][2];
                        row2["fio"] = dt.Rows[i][3];
                        row2["rekvizit"] = rekvizit;
                        dataTable.Rows.Add(row2);
                    }
                }
                catch (Exception)
                {

                }
            }
            return dataTable;
        }

        public DataTable SelectSaldoForSberbank(string localhost, int month, int year, List<string> prefs)
        {
            string connStr = "Server=" + (localhost == "localhost" ? "localhost" : "192.168.1.25") + ";Database=billTlt;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string monthTo = month.ToString("00") + "." + year.ToString("0000");
            string cmdText = "";
            foreach (string pref in prefs)
            {
                if (pref != "bill01")
                    cmdText += " UNION ALL ";
                cmdText += @"SELECT k.pkod, k.fio, 'г. Тольятти ' || ul.ulicareg || ' ' || ul.ulica || ' д. ' || d.idom || ' кв. ' || k.ikvar || ' к. ' || CASE WHEN k.nkvar_n = '' 
                                OR k.nkvar_n  is null THEN '-' ELSE k.nkvar_n END as address, '" + monthTo + @"', sum_outsaldo 
                                FROM " + pref + @"_charge_" + (year - 2000) + @".charge_" +
                                       (pref == "bill02" ? (month - 1).ToString("00") : month.ToString("00")) +
                                @" c 
                                INNER JOIN " + pref + @"_data.kvar k on k.nzp_kvar = c.nzp_kvar 
                                INNER JOIN " + pref + @"_data.dom d on d.nzp_dom = k.nzp_dom
                                INNER JOIN " + pref + @"_data.s_ulica ul on ul.nzp_ul = d.nzp_ul
                                where nzp_serv = 1";
            }
            List<string> kvarParams = new List<string>();
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                da.Fill(dt);
                return dt;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public DataTable SelectSaldoForAvtovazbank2(Int32 year, Int32 month, List<string> prefs)
        {
            string connStr = "Server=192.168.1.25;Database=billTlt;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("pkod");
            dataTable.Columns.Add("qwerty");
            dataTable.Columns.Add("sum_outsaldo");
            dataTable.Columns.Add("fio");
            DataRow row2;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            foreach (string pref in prefs)
            {
                string cmdText = @"SELECT k.pkod, ul.ulicareg || ' ' || ul.ulica || ' д.' || d.ndom || ' кв ' || k.ikvar || ' к ' || CASE WHEN k.nkvar_n = '' 
                                OR k.nkvar_n  is null THEN '-' ELSE k.nkvar_n END as qwerty, sum_outsaldo, k.fio 
                                FROM " + pref + @"_charge_" + (year - 2000).ToString("00") + @".charge_" +
                                       (pref == "bill02" ? (month - 1).ToString("00") : month.ToString("00")) +
                                       @" c 
                                INNER JOIN " + pref + @"_data.kvar k on k.nzp_kvar = c.nzp_kvar 
                                INNER JOIN " + pref + @"_data.dom d on d.nzp_dom = k.nzp_dom
                                INNER JOIN " + pref + @"_data.s_ulica ul on ul.nzp_ul = d.nzp_ul
                                where nzp_serv = 1";
                NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
                NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
                DataTable dt = new DataTable();

                try
                {
                    da.Fill(dt);
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        row2 = dataTable.NewRow();
                        row2["pkod"] = dt.Rows[i][0];
                        row2["qwerty"] = dt.Rows[i][1];
                        row2["sum_outsaldo"] = dt.Rows[i][2];
                        row2["fio"] = dt.Rows[i][3];
                        dataTable.Rows.Add(row2);
                    }
                }
                catch (Exception)
                {

                }
            }
            return dataTable;
        }
    }
}
