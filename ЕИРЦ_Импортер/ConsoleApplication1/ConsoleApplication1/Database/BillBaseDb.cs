using Npgsql;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1.Database
{
    class BillBaseDb
    {
        public string SelectNzpKvarByKvarDom(string database, string nkvar, int nzp_dom)
        {
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "SELECT nzp_kvar FROM bill01_data.kvar k WHERE nkvar = '" + nkvar + "' AND nzp_dom = " + nzp_dom;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                conn.Open();
                da.Fill(dt);
                if (dt.Rows.Count != 1)
                {
                    return "0" + "|Найдено больше или меньше 1-ой улицы";
                }
                else
                {
                    return dt.Rows[0][0].ToString() + "|Найдено";
                }
            }

            catch (Exception e)
            {
                return "0" + "|" + e.ToString();
            }
            finally
            {
                conn.Close();
            }
        }
    }
}
