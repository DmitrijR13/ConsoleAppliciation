using Npgsql;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1.Database
{
    class BillKartDb
    {
        public string SelectNzpKvar(string database, string ulica, string ndom, string nkvar)
        {
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "SELECT nzp_kvar FROM bill01_data.kvar k " +
                             " INNER JOIN bill01_data.dom d on d.nzp_dom = k.nzp_dom " +
                             " INNER JOIN bill01_data.s_ulica ul on ul.nzp_ul = d.nzp_ul " +
                             " WHERE ul.ulica = '" + ulica + "' and ndom = '" + ndom + "' and nkvar = '" + nkvar + "'";
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
                    if (dt.Rows.Count == 0)
                        return "0" + "|Ненайдено ЛС";
                    else
                        return "-1" + "|Найдено больше 1-го ЛС";
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

        public void ClearKart(string database, string nzp_kvar)
        {
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "DELETE FROM bill01_data.gilec where nzp_gil in (SELECT nzp_gil FROM bill01_data.kart where nzp_kvar = " + nzp_kvar + ")";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            //Получаем nzp_prm добавленной нами записи
            cmd.ExecuteNonQuery();

            cmdText = "DELETE FROM bill01_data.kart where nzp_kvar = " + nzp_kvar;
            cmd = new NpgsqlCommand(cmdText, conn);
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        public int InsertGil(string database)
        {
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "INSERT INTO bill01_data.gilec(sogl) VALUES(0) returning nzp_gil";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            //Получаем nzp_prm добавленной нами записи
            int nzp_gil = Convert.ToInt32(cmd.ExecuteScalar());
            conn.Close();
            return nzp_gil;
        }

        public int InsertKart(string database, int nzp_gil, string nzp_kvar, string fam, string ima, string otch, string dat_rog, string gender, int nzp_dok, string serij, string nomer,
            string vid_dat, string vid_mes, string tprp, string dat_sost, string dat_reg, int nzp_rod, string rodstvo,
            string region_op, string okrug_op, string gorod_op, string npunkt_op, string region_ku, string okrug_ku, string gorod_ku, string rem_ku)
        {
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            if (tprp.Length != 1)
            {
                if (tprp.Length != 0)
                    tprp = tprp.Substring(0, 1);
                else
                    tprp = "П";
            }
            int nzp_tkrt = dat_sost == "" ? 1 : 2;
            if (dat_sost == "")
                dat_sost = "01.01.0001";
            if (vid_dat == "")
                vid_dat = "01.01.0001";
            if (region_op.Length > 30)
                region_op = region_op.Substring(0, 30);
            if (okrug_op.Length > 30)
                okrug_op = okrug_op.Substring(0, 30);
            if (gorod_op.Length > 30)
                gorod_op = gorod_op.Substring(0, 30);
            if (npunkt_op.Length > 30)
                npunkt_op = npunkt_op.Substring(0, 30);
            if (okrug_ku.Length > 30)
                okrug_ku = okrug_ku.Substring(0, 30);
            if (gorod_ku.Length > 30)
                gorod_ku = gorod_ku.Substring(0, 30);
            if (rem_ku.Length > 30)
                rem_ku = rem_ku.Substring(0, 30);
            if (region_ku.Length > 30)
                region_ku = region_ku.Substring(0, 30);
            if (nomer.Length > 7)
                nomer = nomer.Substring(0, 7);

            string cmdText = "INSERT INTO bill01_data.kart(nzp_gil, isactual, fam, ima, otch, dat_rog, gender, tprp, nzp_bank, nzp_user, nzp_tkrt, nzp_kvar, nzp_nat, nzp_rod, nzp_dok, serij, " +
                             " nomer, vid_mes, vid_dat, dat_sost, dat_ofor, dat_izm, is_unl, cur_unl, rodstvo, region_op, okrug_op, gorod_op, npunkt_op, region_ku, okrug_ku, gorod_ku, rem_ku) " +
                "VALUES(" + nzp_gil + ", 1, '" + fam + "', '" + ima + "', '" + otch + "', to_date('" + dat_rog + "', 'dd.mm.yyyy'), '" + gender + "', '" + tprp + "', 1, 1, " + nzp_tkrt + ", " + nzp_kvar + ", -1, " + nzp_rod
                + ", " + nzp_dok + ", '" + serij + "', '" + nomer + "', '" + vid_mes + "', to_date('" + vid_dat + "', 'dd.mm.yyyy'),  to_date('" + dat_sost + "', 'dd.mm.yyyy'),  to_date('" + dat_reg
                + "', 'dd.mm.yyyy'), current_date, 1, 1, '" + rodstvo + "','" + region_op + "', '" + okrug_op + "', '" + gorod_op + "', '" + npunkt_op + "', '" + region_ku + "', '"
                + okrug_ku + "', '" + gorod_ku + "', '" + rem_ku + "') returning nzp_kart";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            //Получаем nzp_prm добавленной нами записи
            int nzp_kart = Convert.ToInt32(cmd.ExecuteScalar());
            conn.Close();
            return nzp_kart;
        }

        public void InsertGrgd(int nzpKart)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "INSERT INTO bill01_data.grgd(nzp_kart, nzp_grgd) VALUES(" + nzpKart + ", 1) returning id";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            //Получаем nzp_prm добавленной нами записи
            //int nzp_gil = Convert.ToInt32(cmd.ExecuteScalar());
            conn.Close();
        }

        public int InsertKart(string database, int nzp_gil, string nzp_kvar, string fam, string ima, string otch, string dat_rog, string tprp, string dat_reg, int nzp_rod, string rodstvo)
        {
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            if (tprp.Length != 1)
            {
                if (tprp.Length != 0)
                    tprp = tprp.Substring(0, 1);
                else
                    tprp = "п";
            }
            string cmdText = "";
            if (dat_reg != "")
            {
                cmdText = "INSERT INTO bill01_data.kart(nzp_gil, isactual, fam, ima, otch, dat_rog, tprp, nzp_bank, nzp_user, nzp_tkrt, nzp_kvar, nzp_nat, nzp_rod, dat_sost, dat_ofor, dat_izm, is_unl, cur_unl, rodstvo) " +
                      "VALUES(" + nzp_gil + ", 1, '" + fam + "', '" + ima + "', '" + otch + "', to_date('" + dat_rog + "', 'dd.mm.yyyy'), '" + tprp + "', 1, 1, 1, " + nzp_kvar + ", -1, " + nzp_rod
                      + ", to_date('" + dat_reg + "', 'dd.mm.yyyy'),  to_date('" + dat_reg + "', 'dd.mm.yyyy'), current_date, 1, 1, '" + rodstvo + "') returning nzp_kart";
            }
            else
            {
                cmdText = "INSERT INTO bill01_data.kart(nzp_gil, isactual, fam, ima, otch, dat_rog, tprp, nzp_bank, nzp_user, nzp_tkrt, nzp_kvar, nzp_nat, nzp_rod, dat_izm, is_unl, cur_unl, rodstvo) " +
                      "VALUES(" + nzp_gil + ", 1, '" + fam + "', '" + ima + "', '" + otch + "', to_date('" + dat_rog + "', 'dd.mm.yyyy'), '" + tprp + "', 1, 1, 1, " + nzp_kvar + ", -1, " + nzp_rod
                      + ", current_date, 1, 1, '" + rodstvo + "') returning nzp_kart";
            }

            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            //Получаем nzp_prm добавленной нами записи
            int nzp_kart = Convert.ToInt32(cmd.ExecuteScalar());
            conn.Close();
            return nzp_kart;
        }

    }
}
