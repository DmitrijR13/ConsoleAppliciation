using System;
using System.Collections.Generic;
using System.Data;
using System.Net.Configuration;
using System.Text;
using Npgsql;

namespace ConsoleApplication1
{
    class pg
    {
        private string connStr;
        public pg()
        {
            //connStr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.126.128)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=XE)));User Id = HR; Password = test";
            connStr = "Server=192.168.1.25;Database=billTest;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            //connStr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.5.77)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=orcl)));User Id = MJF; Password = ActaNonVerba";
        }

        public List<string> houses = new List<string>();

        public string InsertPeople5(string gkh_code, Int32 flat, string total_area, string useful_area, string privatized, string residents_count, string fio)
        {
            connStr = "Server=85.140.61.250;Database=gkh_samara;User ID=bars;Password=md5SM3tv;CommandTimeout=180000;";
            //string cmdText = "SELECT gro.id from GKH_REALITY_OBJECT gro inner join GKH_DICT_MUNICIPALITY gdm on GDM.ID = GRO.MUNICIPALITY_ID where GDM.NAME LIKE '%г. Самара, Красноглинский р-н' AND replace(lower(GRO.ADDRESS), ' ','') = replace(lower('" + gkh_code + "'), ' ','')";
            string cmdText = "SELECT id from GKH_REALITY_OBJECT where gkh_code = '" + gkh_code + "'";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            int id;
            conn.Open();
            da.Fill(dt);
            id = Convert.ToInt32(dt.Rows[0][0]);
            int priv = 30;
            if (privatized == "да" || privatized == "Да")
                priv = 10;
            else if (privatized == "Не задано" || privatized == "не задано")
                priv = 30;
            else
                priv = 20;
            if (string.IsNullOrEmpty(total_area) || total_area == " ")
                total_area = "0";
            if (string.IsNullOrEmpty(useful_area) || useful_area == " ")
                useful_area = "0";
            if (string.IsNullOrEmpty(residents_count) || residents_count == " ")
                residents_count = "0";

            if (!houses.Contains(id.ToString()))
            {
                houses.Add(id.ToString());
                string cmdText2 = "DELETE FROM gkh_obj_apartment_info where reality_object_id = " + id;
                NpgsqlCommand cmd2 = new NpgsqlCommand(cmdText2, conn);
                cmd2.ExecuteNonQuery();
            }

            string cmdText1 = "INSERT INTO gkh_obj_apartment_info (object_version, object_create_date, object_edit_date, num_apartment, area_total, count_people, privatized," +
                    "reality_object_id, fio_owner, area_living) VALUES(0, CURRENT_DATE, CURRENT_DATE, '" + flat + "'," + total_area.Replace(',', '.')
                    + "," + residents_count + "," + priv + "," + id + ",'" + fio + "', " + useful_area.Replace(',', '.') + ")";
            NpgsqlCommand cmd1 = new NpgsqlCommand(cmdText1, conn);
            try
            {
                cmd1.ExecuteNonQuery();
                return "ЗАГРУЖЕНО";
            }
            catch (Exception e)
            {
                string err = gkh_code + "||" + flat + "||" + e.Message;
                return err;
            }
            finally
            { conn.Close(); }

        }

        public string SelectPkod(string address, string fio)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            //string connStr = "Server=localhost;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=10000;";
            string cmdText = @"SELECT pkod FROM fbill_data.kvar k
inner join fbill_data.dom d on k.nzp_dom = d.nzp_dom
inner join fbill_data.s_ulica ul on ul.nzp_ul = d.nzp_ul
where upper(ul.ulica) = upper('" + address.Split(',')[0] + "') AND d.ndom = '" + address.Split(',')[1].Split('-')[0]
                         + "' AND k.nkvar = '" + address.Split(',')[1].Split('-')[1] + "' AND replace(upper(fio), ' ','') = replace(upper('" + fio + "'), ' ','')";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                da.Fill(dt);
                if (dt.Rows.Count == 1)
                {
                    return dt.Rows[0][0].ToString();
                }
                else
                {
                    return "00";
                }
            }
            catch
            {
                return "0";
            }
        }

        public List<string> SelectNzpKvar(string address, string nkvar)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            //string connStr = "Server=localhost;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=10000;";
            string cmdText = @"SELECT nzp_kvar, num_ls 
FROM fbill_data.kvar k
inner join fbill_data.dom d on k.nzp_dom = d.nzp_dom
inner join fbill_data.s_ulica ul on ul.nzp_ul = d.nzp_ul
where upper(ul.ulicareg || ' ' || ul.ulica || '  д.' || d.ndom) = upper('" + address + "') AND 'кв. ' || k.nkvar || ' комн. ' ||  k.nkvar_n = '" + nkvar + "'";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                da.Fill(dt);
                if (dt.Rows.Count == 1)
                {
                    return new List<string>()
                    {
                        dt.Rows[0][0].ToString(), 
                        dt.Rows[0][1].ToString()
                    };
                }
                else
                {
                    return null;
                }
            }
            catch (Exception e)
            {
                return null;
            }
        }

        public string SelectPkod2(string address, string nkvar)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            //string connStr = "Server=localhost;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=10000;";
            string cmdText = @"SELECT pkod 
FROM fbill_data.kvar k
inner join fbill_data.dom d on k.nzp_dom = d.nzp_dom
inner join fbill_data.s_ulica ul on ul.nzp_ul = d.nzp_ul
where upper(ul.ulicareg || ' ' || ul.ulica || '  д.' || d.ndom) = upper('" + address + "') AND 'кв. ' || k.nkvar || ' комн. ' ||  k.nkvar_n = '" + nkvar + "'";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                da.Fill(dt);
                if (dt.Rows.Count == 1)
                {
                    return dt.Rows[0][0].ToString();
                }
                else
                {
                    return "00";
                }
            }
            catch (Exception e)
            {
                return "0";
            }
        }

        public String AddCounter(string nzp_kvar, string num_ls, string num_cnt, string dat_uchet, DateTime dat_pay, decimal val_cnt_old, decimal val_cnt_new)
        {
            if (num_cnt == "2.40E+14")
                num_cnt = "2.40001E+14";
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            //string connStr = "Server=localhost;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=10000;";
            string cmdText;
            if (num_cnt == "")
            {
                cmdText = @"SELECT cur_unl, nzp_wp, ist, nzp_counter FROM bill01_data.counters where nzp_serv = 25 AND nzp_kvar = " + nzp_kvar + " " +
                             "AND num_cnt = '" + num_cnt + "' and dat_uchet = to_date('" + dat_uchet + "','dd-mm-yyyy') and val_cnt = " + val_cnt_old;
            }
            else
            {
                cmdText = @"SELECT cur_unl, nzp_wp, ist, nzp_counter FROM bill01_data.counters where nzp_serv = 25 AND nzp_kvar = " + nzp_kvar + " " +
                             "AND dat_uchet = to_date('" + dat_uchet + "','dd-mm-yyyy') and val_cnt = " + val_cnt_old;
            }
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                da.Fill(dt);
                if (dt.Rows.Count == 1)
                {
                    cmdText = @"INSERT INTO bill01_data.counters(nzp_kvar, num_ls, nzp_serv, nzp_cnttype, num_cnt, dat_uchet, val_cnt, is_actual, nzp_user, dat_when, cur_unl, nzp_wp, ist, nzp_counter) 
                              VALUES(" + nzp_kvar + ", " + num_ls + ", 25, 17, " + num_cnt + ", " + dat_pay + ", " + val_cnt_new + ", 1, 1, current_date, " + dt.Rows[0][0].ToString()
                                       + ", " + dt.Rows[0][1].ToString() + ", " + dt.Rows[0][2].ToString() + ", " + dt.Rows[0][3].ToString() + ")";
                    cmd = new NpgsqlCommand(cmdText, conn);
                    conn.Open();
                    cmd.ExecuteNonQuery();
                    conn.Close();
                    return "Success";
                }
                else if (dt.Rows.Count > 2)
                {
                    conn.Open();
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        cmdText = @"INSERT INTO bill01_data.counters(nzp_kvar, num_ls, nzp_serv, nzp_cnttype, num_cnt, dat_uchet, val_cnt, is_actual, nzp_user, dat_when, cur_unl, nzp_wp, ist, nzp_counter) 
                              VALUES(" + nzp_kvar + ", " + num_ls + ", 25, 17, " + num_cnt + ", " + dat_pay + ", " + val_cnt_new + ", 1, 1, current_date, " + dt.Rows[i][0].ToString()
                                       + ", " + dt.Rows[i][1].ToString() + ", " + dt.Rows[i][2].ToString() + ", " + dt.Rows[i][3].ToString() + ")";
                        cmd = new NpgsqlCommand(cmdText, conn);
                        cmd.ExecuteNonQuery();
                    }
                    conn.Close();
                    return "Success";
                }
                else
                {
                    return "Не найдено не одного счетчика по входным параметрам";
                }
            }
            catch (Exception e)
            {
                return "Ошибка";
            }
        }

        public int UpdateDatePaid(string pkod, string dateWrong, string sum, string dateRirgt, string datePaid)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";

            string cmdText = @"SELECT * FROM fbill_fin_15.pack_ls " +
                " where pkod = " + pkod + " AND dat_vvod = to_date('" + dateWrong + "','dd.mm.yyyy') and g_sum_ls = " + sum +
                " and dat_month = to_date('01." + datePaid.Substring(0, 2) + ".20" + datePaid.Substring(2) + "','dd.mm.yyyy')";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                da.Fill(dt);
                if (dt.Rows.Count == 1)
                {
                    cmdText = "UPDATE fbill_fin_15.pack_ls SET dat_vvod = to_date('" + dateRirgt + "','dd.mm.yyyy') " +
                        " where pkod = " + pkod + " AND dat_vvod = to_date('" + dateWrong + "','dd.mm.yyyy') and g_sum_ls = " + sum +
                        " and dat_month = to_date('01." + datePaid.Substring(0, 2) + ".20" + datePaid.Substring(2) + "','dd.mm.yyyy')";
                    cmd = new NpgsqlCommand(cmdText, conn);
                    conn.Open();
                    int i = cmd.ExecuteNonQuery();
                    conn.Close();
                    return i;
                }
                else if (dt.Rows.Count == 0)
                {
                    return 0;
                }
                else
                {
                    return 2;
                }
            }
            catch
            {
                return 0;
            }

        }

        public string InsertKvar(String num_ls, String fio, String nkvar, Int32 ikvar, String nkvar_n)
        {
            string connStr = "Server=192.168.1.25;Database=billTlt;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            try
            {
                //Добавляем строку в pref_kernel.prm_name
                string cmdText = @"INSERT INTO fbill_data.kvar(
            nzp_area, nzp_geu, nzp_dom, nkvar, nkvar_n, num_ls, 
            fio, ikvar,
            typek, pref, is_open, nzp_wp, area_code)
    VALUES (2, 2, 7155106, '" + nkvar + "', '" + nkvar_n + "', " + num_ls + @", 
            '" + fio + @"', " + ikvar + @",
            1, 'bill01', 1, 25, 40302) returning nzp_kvar";
                NpgsqlConnection conn = new NpgsqlConnection(connStr);
                NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
                conn.Open();
                //Получаем nzp_prm добавленной нами записи
                int nzp_kvar = Convert.ToInt32(cmd.ExecuteScalar());

                //опускаем в банки ниже    
                cmdText = @" INSERT INTO bill01_data.kvar
 SELECT nzp_kvar, nzp_area, nzp_geu, nzp_dom, nkvar, nkvar_n, num_ls, 
            porch, phone, dat_notp_s, dat_notp_po, fio, ikvar, uch, gil_s, 
            remark, typek, pkod, pkod10 FROM  fbill_data.kvar where nzp_kvar = " + nzp_kvar;
                cmd = new NpgsqlCommand(cmdText, conn);
                cmd.ExecuteNonQuery();
                conn.Close();
                return "Success";
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
            
        }

        public void InsertTarif(string database, int nzp_serv, int nzp_measure, string name_frm, string name_prm, string dat_s, string dat_po,
            string val_prm, List<string> prefs, string is_actual, string nzp_frm_typ, string nzp_frm_typrs, string nzp_prm_rash, string nzp_prm_rash1)
        {
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            //Добавляем строку в pref_kernel.prm_name
            string cmdText = "INSERT INTO fbill_kernel.prm_name(name_prm, old_field, type_prm, prm_num, digits_, is_day_uchet) values('"+
                name_prm + "', 0, 'float', 5, 4, 0) returning nzp_prm";
            NpgsqlConnection conn= new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            //Получаем nzp_prm добавленной нами записи
            int nzp_prm = Convert.ToInt32(cmd.ExecuteScalar());
            Console.WriteLine("nzp_prm = " + nzp_prm);
            //опускаем в банки ниже
            foreach (string pref in prefs)
            {
                cmdText = "INSERT INTO " + pref + "_kernel.prm_name(nzp_prm, name_prm, old_field, type_prm, prm_num, digits_, is_day_uchet) values(" + nzp_prm + ", '" +
                name_prm + "', 0, 'float', 5, 4, 0)";
                cmd = new NpgsqlCommand(cmdText, conn);
                cmd.ExecuteNonQuery();
            }
            //Добавляем строку в pref_kernel.formuls
            cmdText = "INSERT INTO fbill_kernel.formuls(name_frm, dat_s, dat_po, nzp_measure, is_device) values('"
                + name_frm + "', to_date('01.01.2000', 'dd.mm.yyyy'), to_date('01.01.3000', 'dd.mm.yyyy'), " + nzp_measure + ", 0) returning nzp_frm";
            cmd = new NpgsqlCommand(cmdText, conn);
            //Получаем nzp_frm добавленной нами записи
            int nzp_frm = Convert.ToInt32(cmd.ExecuteScalar());
            Console.WriteLine("nzp_frm = " + nzp_frm);
            foreach (string pref in prefs)
            {
                cmdText = "INSERT INTO " + pref + "_kernel.formuls(nzp_frm, name_frm, dat_s, dat_po, nzp_measure, is_device) values(" + nzp_frm + ",'"
                + name_frm + "', to_date('01.01.2000', 'dd.mm.yyyy'), to_date('01.01.3000', 'dd.mm.yyyy'), " + nzp_measure + ", 0)";
                cmd = new NpgsqlCommand(cmdText, conn);
                cmd.ExecuteNonQuery();
            }
            //Добавляем строку в pref_kernel.formuls_opis
            cmdText = "insert into fbill_kernel.formuls_opis " +
                    "(nzp_frm, nzp_frm_kod, nzp_frm_typ, nzp_prm_tarif_ls, nzp_prm_tarif_lsp, nzp_prm_tarif_dm, nzp_prm_tarif_su, nzp_prm_tarif_bd, nzp_frm_typrs, nzp_prm_rash, nzp_prm_rash1, nzp_prm_rash2) " +
                     "values (" + nzp_frm + ", " + nzp_frm + ", " + nzp_frm_typ + ", 0, 0, 0, 0, " + nzp_prm + ", " + nzp_frm_typrs + ", " + nzp_prm_rash + ", " + nzp_prm_rash1 + ", 0) returning nzp_ops";
            cmd = new NpgsqlCommand(cmdText, conn);
            //Получаем nzp_frm_opis добавленной нами записи
            int nzp_frm_opis = Convert.ToInt32(cmd.ExecuteScalar());
            foreach (string pref in prefs)
            {
                cmdText = "insert into " + pref + "_kernel.formuls_opis " +
                    "(nzp_ops, nzp_frm, nzp_frm_kod, nzp_frm_typ, nzp_prm_tarif_ls, nzp_prm_tarif_lsp, nzp_prm_tarif_dm, nzp_prm_tarif_su, nzp_prm_tarif_bd, nzp_frm_typrs, nzp_prm_rash, nzp_prm_rash1, nzp_prm_rash2) " +
                     "values (" + nzp_frm_opis + ", " + nzp_frm + ", " + nzp_frm + ", " + nzp_frm_typ + ", 0, 0, 0, 0, " + nzp_prm + ", " + nzp_frm_typrs + ", " + nzp_prm_rash + ", " + nzp_prm_rash1 + ", 0)";
                cmd = new NpgsqlCommand(cmdText, conn);
                cmd.ExecuteNonQuery();
            }

            //Добавляем строку в pref_kernel.prm_tarifs
            cmdText = "INSERT INTO fbill_kernel.prm_tarifs(nzp_serv, nzp_frm, nzp_prm, is_edit, nzp_user, dat_when) values("+
                nzp_serv + ", " + nzp_frm + ", " + nzp_prm + ", 1 , 1, to_date('01.06.2014', 'dd.mm.yyyy')) returning nzp_key";
            cmd = new NpgsqlCommand(cmdText, conn);
            //Получаем nzp_key добавленной нами записи
            int nzp_key = Convert.ToInt32(cmd.ExecuteScalar());
            Console.WriteLine("nzp_key = " + nzp_key);
            foreach (string pref in prefs)
            {
                cmdText = "INSERT INTO " + pref + "_kernel.prm_tarifs(nzp_key, nzp_serv, nzp_frm, nzp_prm, is_edit, nzp_user, dat_when) values(" +
                nzp_key + ", " + nzp_serv + ", " + nzp_frm + ", " + nzp_prm + ", 1 , 1, to_date('01.06.2014', 'dd.mm.yyyy'))";
                cmd = new NpgsqlCommand(cmdText, conn);
                cmd.ExecuteNonQuery();
            }

            //Добавляем строку в pref_bill01_data.prm_5
            cmdText = "INSERT INTO fbill_data.prm_5(nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user,dat_when) "+
                "VALUES (0, " + nzp_prm + ",  to_date('" + dat_s + "', 'dd.mm.yyyy'),  to_date('" + dat_po + "', 'dd.mm.yyyy'), '" + val_prm + "', "+is_actual+", 1, to_date('01.06.2014', 'dd.mm.yyyy')) returning nzp_key";
            cmd = new NpgsqlCommand(cmdText, conn);
            //Получаем nzp_key добавленной нами записи
            int nzp_key2 = Convert.ToInt32(cmd.ExecuteScalar());
            foreach (string pref in prefs)
            {
                cmdText = "INSERT INTO " + pref + "_data.prm_5(nzp_key, nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user,dat_when) " +
                "VALUES (" + nzp_key2 + ", 0, " + nzp_prm + ",  to_date('" + dat_s + "', 'dd.mm.yyyy'),  to_date('" + dat_po + "', 'dd.mm.yyyy'), '" + 
                val_prm + "', "+is_actual+", 1, to_date('01.06.2014', 'dd.mm.yyyy'))";
                cmd = new NpgsqlCommand(cmdText, conn);
                cmd.ExecuteNonQuery();
            }

            //Добавляем строку в fbill_kernel.prm_frm
            cmdText = "INSERT INTO fbill_kernel.prm_frm(nzp_frm, frm_calc, \"order\", is_prm, operation, nzp_prm) VALUES(" + nzp_frm + ", 99, 1, 1, 'FLD', " + nzp_prm + ") returning nzp_pf";
            cmd = new NpgsqlCommand(cmdText, conn);
            int nzp_pf = Convert.ToInt32(cmd.ExecuteScalar());
            Console.WriteLine("nzp_pf = " + nzp_pf);
            foreach (string pref in prefs)
            {
                cmdText = "INSERT INTO " + pref + "_kernel.prm_frm(nzp_pf, nzp_frm, frm_calc, \"order\", is_prm, operation, nzp_prm) VALUES("+nzp_pf+"," + nzp_frm + ", 99, 1, 1, 'FLD', " + nzp_prm + ")";
                cmd = new NpgsqlCommand(cmdText, conn);
                cmd.ExecuteNonQuery();
            }

            conn.Close();
        }

        public void InsertNewLS(int num_ls, int flat_num)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            //Добавляем строку в pref_kernel.prm_name
            string cmdText = "INSERT INTO fbill_data.kvar(nzp_area, nzp_geu, nzp_dom, nkvar, num_ls, fio, ikvar, typek, pref, is_open, nzp_wp, area_code) " +
                " VALUES(2, 2, 7155104, '" + flat_num + "', " + num_ls + ", '', " + flat_num + ", 1, 'bill01', 1, 25, 40102) returning nzp_kvar";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            //Получаем nzp_prm добавленной нами записи
            int nzp_kvar = Convert.ToInt32(cmd.ExecuteScalar());
            Console.WriteLine("nzp_kvar = " + nzp_kvar);
            //опускаем в банки ниже
            cmdText = "INSERT INTO bill01_data.kvar(nzp_kvar, nzp_area, nzp_geu, nzp_dom, nkvar, nkvar_n, num_ls, fio, ikvar, typek) " +
" VALUES(" + nzp_kvar + ", 2, 2, 7155104, '" + flat_num + "', '-', " + num_ls + ", '', " + flat_num + ", 1)";
            cmd = new NpgsqlCommand(cmdText, conn);
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        public int InsertPrmName(string name_prm, List<string> prefs)
        {
            //Добавляем строку в pref_kernel.prm_name
            string cmdText = "INSERT INTO fbill_kernel.prm_name(name_prm, old_field, type_prm, prm_num) values('" +
                name_prm + "', 0, 'int', 13) returning nzp_prm";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            //Получаем nzp_prm добавленной нами записи
            int nzp_prm = Convert.ToInt32(cmd.ExecuteScalar());
            //опускаем в банки ниже
            foreach (string pref in prefs)
            {
                cmdText = "INSERT INTO " + pref + "_kernel.prm_name(nzp_prm, name_prm, old_field, type_prm, prm_num) values(" + nzp_prm + ", '" +
                name_prm + "', 0, 'int', 13)";
                cmd = new NpgsqlCommand(cmdText, conn);
                cmd.ExecuteNonQuery();
            }
            conn.Close();
            return nzp_prm;
        }

        public int InsertResolution(string name_short, string name_res, List<string> prefs)
        {
            //Добавляем строку в pref_kernel.prm_name
            string cmdText = "INSERT INTO fbill_kernel.resolution(name_short, name_res, is_readonly) values('" +
                name_short + "', '" + name_res + "', 1) returning nzp_res";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            //Получаем nzp_prm добавленной нами записи
            int nzp_res = Convert.ToInt32(cmd.ExecuteScalar());
            //опускаем в банки ниже
            foreach (string pref in prefs)
            {
                cmdText = "INSERT INTO " + pref + "_kernel.resolution(nzp_res, name_short, name_res, is_readonly) values(" + nzp_res + ", '" +
                name_short + "', '" + name_res + "', 1)";
                cmd = new NpgsqlCommand(cmdText, conn);
                cmd.ExecuteNonQuery();
            }
            conn.Close();
            return nzp_res;
        }

        public void InsertResX(int nzp_res, int nzp_x, string name_x, List<string> prefs)
        {
            //Добавляем строку в pref_kernel.prm_name
            string cmdText = "INSERT INTO fbill_kernel.res_x(nzp_res, nzp_x, name_x) values(" +
                nzp_res + ", " + nzp_x + ", '" + name_x + "')";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            //опускаем в банки ниже
            foreach (string pref in prefs)
            {
                cmdText = "INSERT INTO " + pref + "_kernel.res_x(nzp_res, nzp_x, name_x) values(" +
                nzp_res + ", " + nzp_x + ", '" + name_x + "')";
                cmd = new NpgsqlCommand(cmdText, conn);
                cmd.ExecuteNonQuery();
            }
            conn.Close();
        }

        public void InsertResY(int nzp_res, int nzp_y, string name_y, List<string> prefs)
        {
            //Добавляем строку в pref_kernel.prm_name
            string cmdText = "INSERT INTO fbill_kernel.res_y(nzp_res, nzp_y, name_y) values(" +
                nzp_res + ", " + nzp_y + ", '" + name_y + "')";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            //опускаем в банки ниже
            foreach (string pref in prefs)
            {
                cmdText = "INSERT INTO " + pref + "_kernel.res_y(nzp_res, nzp_y, name_y) values(" +
                nzp_res + ", " + nzp_y + ", '" + name_y + "')";
                cmd = new NpgsqlCommand(cmdText, conn);
                cmd.ExecuteNonQuery();
            }
            conn.Close();
        }

        public void InsertResValue(int nzp_res, int nzp_y, int nzp_x, string value, List<string> prefs)
        {
            //Добавляем строку в pref_kernel.prm_name
            string cmdText = "INSERT INTO fbill_kernel.res_values(nzp_res, nzp_y, nzp_x, value) values(" +
                nzp_res + ", " + nzp_y + ", " + nzp_x + ", '" + value + "')";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            //опускаем в банки ниже
            foreach (string pref in prefs)
            {
                cmdText = "INSERT INTO " + pref + "_kernel.res_values(nzp_res, nzp_y, nzp_x, value) values(" +
                nzp_res + ", " + nzp_y + ", " + nzp_x + ", '" + value + "')";
                cmd = new NpgsqlCommand(cmdText, conn);
                cmd.ExecuteNonQuery();
            }
            conn.Close();
        }

        public void InsertPrm13(int nzp_prm, string dat_s, string dat_po, int val_prm, string is_actual, string dat_when, List<string> prefs)
        {
            //Добавляем строку в pref_kernel.prm_name
            string cmdText = "INSERT INTO fbill_data.prm_13(nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user, dat_when) values(0, "+
                nzp_prm + ", to_date('" + dat_s + "', 'dd.mm.yyyy'), to_date('" + dat_po + "', 'dd.mm.yyyy'), '" + val_prm + "', " + is_actual + ", 1, to_date('" + dat_when + "', 'dd.mm.yyyy')) returning nzp_key";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            //Получаем nzp_prm добавленной нами записи
            int nzp_key = Convert.ToInt32(cmd.ExecuteScalar());
            //опускаем в банки ниже
            foreach (string pref in prefs)
            {
                cmdText = "INSERT INTO " + pref + "_data.prm_13(nzp_key, nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user, dat_when) values(" + nzp_key + ", 0, " +
                nzp_prm + ", to_date('" + dat_s + "', 'dd.mm.yyyy'), to_date('" + dat_po + "', 'dd.mm.yyyy'), '" + val_prm + "', " + is_actual + ", 1, to_date('" + dat_when + "', 'dd.mm.yyyy'))";
                cmd = new NpgsqlCommand(cmdText, conn);
                cmd.ExecuteNonQuery();
            }
            conn.Close();
        }

        public void InsertResValues2(string nzp_res, string nzp_y, string nzp_x, string value, string id, string pref)
        {
            //Добавляем строку в pref_kernel.prm_name
            string cmdText = "INSERT INTO " + pref + "_kernel.res_values(nzp_res, nzp_y, nzp_x, value, id) values(" +
                nzp_res + ", " + nzp_y + ", " + nzp_x + ", '" + value + "', " + id + ")";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        public void InsertResX2(string nzp_res, string nzp_x, string name_x, string id, string pref)
        {
            string cmdText = "INSERT INTO " + pref + "_kernel.res_x(nzp_res, nzp_x, name_x, id) values(" +
                nzp_res + ", " + nzp_x + ", '" + name_x + "', " + id + ")";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        public void InsertResY2(string nzp_res, string nzp_y, string name_y, string id, string pref)
        {
            string cmdText = "INSERT INTO " + pref + "_kernel.res_y(nzp_res, nzp_y, name_y, id) values(" +
                nzp_res + ", " + nzp_y + ", '" + name_y + "', " + id + ")";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        public void InsertResolution2(string nzp_res, string name_short, string name_res, string pref)
        {
            //Добавляем строку в pref_kernel.prm_name
            string cmdText = "INSERT INTO " + pref + "_kernel.resolution(nzp_res, name_short, name_res, is_readonly) values(" + nzp_res + ",'" +
                name_short + "', '" + name_res + "', 1)";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        public void InsertPrmName2(string nzp_prm, string name_prm, string pref)
        {
            //Добавляем строку в pref_kernel.prm_name
            string cmdText = "INSERT INTO " + pref + "_kernel.prm_name(nzp_prm, name_prm, old_field, type_prm, prm_num) values(" + nzp_prm + ",'" +
                name_prm + "', 0, 'int', 13) returning nzp_prm";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        public string InsertPayment(string doc_number, string doc_date, string amount, string kbk, string recipient, string gos_contract_doc_number, string source_subkesr, string kosgu, string code)
        {
            string connStr = "Server=localhost;Database=sobits_str;User ID=sobits_str;Password=68912022;CommandTimeout=180000;";
            string cmdText = "SELECT id FROM mosks_object_aip where code = 'ОКС - " + code + "'";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            int object_aip_id;
            try
            {
                conn.Open();
                da.Fill(dt);
                object_aip_id = Convert.ToInt32(dt.Rows[0][0]);
                cmdText = "SELECT id FROM mosks_gos_contract where doc_number = '" + gos_contract_doc_number + "'";
                cmd = new NpgsqlCommand(cmdText, conn);
                da = new NpgsqlDataAdapter(cmd);
                dt = new DataTable();
                int gos_contract_id;
                da.Fill(dt);
                gos_contract_id = Convert.ToInt32(dt.Rows[0][0]);
                cmdText = "SELECT id FROM mosks_gos_contract_rec where gos_contract_id = " + gos_contract_id;
                cmd = new NpgsqlCommand(cmdText, conn);
                da = new NpgsqlDataAdapter(cmd);
                dt = new DataTable();
                int gos_contract_rec_id;
                da.Fill(dt);
                gos_contract_rec_id = Convert.ToInt32(dt.Rows[0][0]);
                cmdText = @"INSERT INTO mosks_object_aip_payment(object_version, object_create_date, object_edit_date, doc_number, doc_date, object_aip_id, gos_contract_rec_id, date_accept, amount, payer, recipient, repaid, returned, payment_type,
  is_avance, is_deleted_from_bo, kosgu,source_subkesr, is_writ_of_execution, is_accruals_and_expense_report)
  values(0, CURRENT_DATE, CURRENT_DATE, '" + doc_number + "', to_date('" + doc_date + "','dd.mm.yyyy'), " + object_aip_id + ", " + gos_contract_rec_id +
                                              ", '-infinity', " + amount + ", 'Департамент строительства и архитектуры г.о. Самара', '" +
                                              recipient + "', 0, 0, 10, false, false, " + kosgu + ", " + source_subkesr + ", false, false)  returning id";
                Console.WriteLine(cmdText);
                cmd = new NpgsqlCommand(cmdText, conn);
                //Получаем nzp_frm добавленной нами записи
                int object_aip_payment_id = Convert.ToInt32(cmd.ExecuteScalar());
                cmdText = @"INSERT INTO mosks_object_aip_payment_rec(object_version, object_create_date, object_edit_date, object_aip_payment_id, gos_contract_rec_id, amount, kbk, is_deleted_from_bo) 
                        values(0, CURRENT_DATE, CURRENT_DATE, " + object_aip_payment_id + ", " + gos_contract_rec_id + ", " + amount + ", '" + kbk + "', false)";
                Console.WriteLine(cmdText);
                cmd = new NpgsqlCommand(cmdText, conn);
                cmd.ExecuteNonQuery();
                return "ЗАГРУЖЕНО";
            }
            catch (Exception e)
            {
                string err = e.Message;
                return err + " = " + doc_number;
            }
            finally
            { 
                conn.Close(); 
            }
        }

        public string InsertRecipientID(string part1, string part2)
        {
            string connStr = "Server=localhost;Database=sobits_str;User ID=sobits_str;Password=68912022;CommandTimeout=180000;";
            string cmdText = "SELECT id, name FROM eas_organization where replace(upper(name), ' ','') like replace(upper('%"+part1+"%'), ' ','') and replace(upper(name), ' ','') like replace(upper('%"+part2+"%'), ' ','')";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            int object_aip_id;
            string name;
            try
            {
                conn.Open();
                da.Fill(dt);
                if (dt.Rows.Count == 0)
                {
                    return "НЕ ЗАГРУЖЕНО|0";
                }
                else if (dt.Rows.Count == 1)
                {
                    object_aip_id = Convert.ToInt32(dt.Rows[0][0]);
                    name = Convert.ToString(dt.Rows[0][1]);
                    return "ЗАГРУЖЕНО|" + object_aip_id + "|" + name;
                }
                else
                {
                    return "НЕ ЗАГРУЖЕНО|2";
                }
            }
            catch (Exception e)
            {
                string err = e.Message;
                return "НЕ ЗАГРУЖЕНО|" + err;
            }
            finally
            {
                conn.Close();
            }
        }

        public int InsertRecipientID(string name)
        {
            string connStr = "Server=localhost;Database=sobits_str;User ID=sobits_str;Password=68912022;CommandTimeout=180000;";
            string cmdText = "SELECT id, name FROM eas_organization where replace(upper(name), ' ','') like replace(upper('%" + name + "%'), ' ','')";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            int object_aip_id;
            try
            {
                da.Fill(dt);
                if (dt.Rows.Count == 0)
                {
                    return 0;
                }
                else if (dt.Rows.Count == 1)
                {
                    object_aip_id = Convert.ToInt32(dt.Rows[0][0]);
                    return object_aip_id;
                }
                else
                {
                    return 0;
                }
            }
            catch (Exception e)
            {
                return 0;
            }
            finally
            {
                conn.Close();
            }
        }

        public int GetObjectId(string name)
        {
            string connStr = "Server=localhost;Database=sobits_str;User ID=sobits_str;Password=68912022;CommandTimeout=180000;";
            string cmdText = "select id from mosks_object_aip where replace(replace(upper(name), ' ',''), '\"','') like replace(replace(upper('%"+name+"%'), ' ',''), '\"','')";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            int object_aip_id;
            try
            {
                da.Fill(dt);
                if (dt.Rows.Count == 0)
                {
                    try
                    {
                        string name1;
                        name1 = name.Substring(0,90);
                        cmdText = "select id from mosks_object_aip where replace(replace(upper(name), ' ',''), '\"','') like replace(replace(upper('%" + name1 + "%'), ' ',''), '\"','')";
                        conn = new NpgsqlConnection(connStr);
                        cmd = new NpgsqlCommand(cmdText, conn);
                        da = new NpgsqlDataAdapter(cmd);
                        dt = new DataTable();
                        da.Fill(dt);
                        if (dt.Rows.Count == 0)
                        {
                            try
                            {
                                string name2;
                                name2 = name.Substring(85, 90);
                                cmdText = "select id from mosks_object_aip where replace(replace(upper(name), ' ',''), '\"','') like replace(replace(upper('%" + name2 + "%'), ' ',''), '\"','')";
                                conn = new NpgsqlConnection(connStr);
                                cmd = new NpgsqlCommand(cmdText, conn);
                                da = new NpgsqlDataAdapter(cmd);
                                dt = new DataTable();
                                da.Fill(dt);
                                if (dt.Rows.Count == 0)
                                {
                                    return 0;
                                }
                                else if (dt.Rows.Count == 1)
                                {
                                    object_aip_id = Convert.ToInt32(dt.Rows[0][0]);
                                    return object_aip_id;
                                }
                                else
                                {
                                    return 2;
                                }
                            }
                            catch
                            {
                                return 0;
                            }
                        }
                        else if (dt.Rows.Count == 1)
                        {
                            object_aip_id = Convert.ToInt32(dt.Rows[0][0]);
                            return object_aip_id;
                        }
                        else
                        {
                            return 2;
                        }
                    }
                    catch
                    {
                        return 0;
                    }
                }
                else if (dt.Rows.Count == 1)
                {
                    object_aip_id = Convert.ToInt32(dt.Rows[0][0]);
                    return object_aip_id;
                }
                else
                {
                    return 2;
                }
            }
            catch (Exception e)
            {
                return 1;
            }
            finally
            {
                conn.Close();
            }
        }

        public int InsertRecipient(string name)
        {
            string connStr = "Server=localhost;Database=sobits_str;User ID=sobits_str;Password=68912022;CommandTimeout=180000;";
            string cmdText = "INSERT INTO eas_organization(object_version, object_create_date, object_edit_date , name, short_name, organization_form) " +
                "VALUES(0, CURRENT_DATE, CURRENT_DATE, '" + name + "', '" + name + "', 0) returning id";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            int recipient_id = 0;
            conn.Open();
            try
            {
                recipient_id = Convert.ToInt32(cmd.ExecuteScalar());
                return recipient_id;
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
                recipient_id = 0;
                return recipient_id;
            }
            finally
            {
                conn.Close();
            }
        }

        public void UpdateSaldo(decimal serv25, decimal serv515, decimal serv1, int nzp_kvar, decimal c_calc25, decimal c_calc515)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "UPDATE bill01_charge_15.charge_01 SET tarif = 3.17, rsum_tarif = " + serv25 +", sum_tarif = "+serv25+", sum_real = "+serv25+
                ", sum_charge = " + serv25 + ", sum_outsaldo = " + serv25 + ", tarif_f = 3.17, sum_tarif_f = " + serv25 + ", gsum_tarif = " + serv25 + 
                " where nzp_serv = 25 AND nzp_kvar = " + nzp_kvar;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            try
            {
                cmd.ExecuteNonQuery();
                conn.Close();
                cmdText = "UPDATE bill01_charge_15.charge_01 SET tarif = 3.17, rsum_tarif = " + serv515 + ", sum_tarif = " + serv515 + ", sum_real = " + serv515 +
                    ", sum_charge = " + serv515 + ", sum_outsaldo = " + serv515 + ", tarif_f = 3.17, sum_tarif_f = " + serv515 + ", gsum_tarif = " + serv515 +
                    " where nzp_serv = 515 AND nzp_kvar = " + nzp_kvar;
                conn = new NpgsqlConnection(connStr);
                cmd = new NpgsqlCommand(cmdText, conn);
                try
                {
                    conn.Open();
                    cmd.ExecuteNonQuery();
                    conn.Close();
                    cmdText = "UPDATE bill01_charge_15.charge_01 SET rsum_tarif = rsum_tarif + " + serv1 + ", sum_tarif = sum_tarif + " + serv1 + ", sum_real = sum_real + " + serv1 +
                   ", sum_charge = sum_charge + " + serv1 + ", sum_outsaldo = sum_outsaldo + " + serv1 + ", sum_tarif_f = sum_tarif_f + " + serv1 + ", gsum_tarif = gsum_tarif + " + serv1 +
                   " where nzp_serv = 1 AND nzp_kvar = " + nzp_kvar;
                    conn = new NpgsqlConnection(connStr);
                    cmd = new NpgsqlCommand(cmdText, conn);
                    try
                    {
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                        cmdText = "UPDATE bill01_charge_15.charge_02 SET sum_insaldo = sum_insaldo + " + serv515 +
                  " where nzp_serv = 515 AND nzp_kvar = " + nzp_kvar;
                        conn = new NpgsqlConnection(connStr);
                        cmd = new NpgsqlCommand(cmdText, conn);
                        try
                        {
                            conn.Open();
                            cmd.ExecuteNonQuery();
                            conn.Close();
                            cmdText = "UPDATE bill01_charge_15.charge_02 SET sum_insaldo = sum_insaldo + " + serv25 +
                " where nzp_serv = 25 AND nzp_kvar = " + nzp_kvar;
                            conn = new NpgsqlConnection(connStr);
                            cmd = new NpgsqlCommand(cmdText, conn);
                            try
                            {
                                conn.Open();
                                cmd.ExecuteNonQuery();
                                conn.Close();
                                cmdText = "UPDATE bill01_charge_15.charge_01 SET c_calc = " + c_calc25 +
                                " where nzp_serv = 25 AND nzp_kvar = " + nzp_kvar;
                                conn = new NpgsqlConnection(connStr);
                                cmd = new NpgsqlCommand(cmdText, conn);
                                try
                                {
                                    conn.Open();
                                    cmd.ExecuteNonQuery();
                                    conn.Close();
                                    cmdText = "UPDATE bill01_charge_15.charge_01 SET c_calc = " + c_calc515 +
                                    " where nzp_serv = 515 AND nzp_kvar = " + nzp_kvar;
                                    conn = new NpgsqlConnection(connStr);
                                    cmd = new NpgsqlCommand(cmdText, conn);
                                    try
                                    {
                                        conn.Open();
                                        cmd.ExecuteNonQuery();
                                        conn.Close();
                                        cmdText = "UPDATE bill01_charge_15.charge_01 SET c_sn = " + c_calc25 +
                                        " where nzp_serv = 25 AND nzp_kvar = " + nzp_kvar;
                                        conn = new NpgsqlConnection(connStr);
                                        cmd = new NpgsqlCommand(cmdText, conn);
                                        try
                                        {
                                            conn.Open();
                                            cmd.ExecuteNonQuery();
                                            conn.Close();
                                        }
                                        catch { }
                                    }
                                    catch { }
                                }
                                catch
                                {
                                }
                            }
                            catch { }
                        }
                        catch { }
                    }
                    catch (Exception ex)
                    { }

                }
                catch (Exception ex)
                {
                    string atr = "";
                    Console.WriteLine(ex.ToString());
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                conn.Close();
            }


        }

        public int InsertGosContract(string docNumber, DateTime docDate, int recipientId, string amount, DateTime termsOfObligation)
        {
            string connStr = "Server=localhost;Database=sobits_str;User ID=sobits_str;Password=68912022;CommandTimeout=180000;";

            string cmdText = "SELECT * FROM mosks_gos_contract where doc_number = '"+docNumber+"' AND doc_date = to_date('"+docDate.ToShortDateString()+"', 'dd.mm.yyyy')";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            if (dt.Rows.Count != 0)
            {
                conn.Close();
                return 0;
            }
            StringBuilder sbIns = new StringBuilder();
            StringBuilder sbVal = new StringBuilder();
            sbIns.Append("INSERT INTO mosks_gos_contract(object_version, object_create_date, object_edit_date, doc_number, doc_date, customer_id, recipient_id , amount, ");
            sbVal.Append("  VALUES(0, CURRENT_DATE, CURRENT_DATE, '" + docNumber + "', to_date('" + docDate.ToShortDateString() + "', 'dd.mm.yyyy'), 14, "+recipientId+", "+amount+", ");
            if (termsOfObligation != new DateTime(1111, 1, 1))
            {
                sbIns.Append("terms_of_obligations, ");
                sbVal.Append("to_date('" + termsOfObligation.ToShortDateString() + "', 'dd.mm.yyyy'), ");
            }
            sbIns.Append("document_type_id)");
            sbVal.Append("130)");
            cmdText = sbIns.ToString() + sbVal.ToString() + " returning id";
            Console.WriteLine(cmdText);
            cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            int gos_contract_id;
            try
            {
                gos_contract_id = Convert.ToInt32(cmd.ExecuteScalar());
                return gos_contract_id;
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
                gos_contract_id = 1;
                return gos_contract_id;
            }
            finally
            {
                conn.Close();
            }
        }

        public int InsertGosContractRec(int gosContractId, int objecAipId, string amount)
        {
            string connStr = "Server=localhost;Database=sobits_str;User ID=sobits_str;Password=68912022;CommandTimeout=180000;";
            StringBuilder sbIns = new StringBuilder();
            StringBuilder sbVal = new StringBuilder();
            sbIns.Append("INSERT INTO mosks_gos_contract_rec(object_version, object_create_date, object_edit_date, gos_contract_id, ");
            sbVal.Append("  VALUES(0, CURRENT_DATE, CURRENT_DATE, " + gosContractId + ", ");
            if (objecAipId != 0)
            {
                sbIns.Append("object_aip_id, ");
                sbVal.Append(objecAipId + ", ");
            }
            sbIns.Append("amount)");
            sbVal.Append(amount+ ")");
            string cmdText = sbIns.ToString() + sbVal.ToString() + " returning id";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            int gos_contract_rec_id;
            try
            {
                gos_contract_rec_id = Convert.ToInt32(cmd.ExecuteScalar());
                return gos_contract_rec_id;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                gos_contract_rec_id = 0;
                return gos_contract_rec_id;
            }
            finally
            {
                conn.Close();
            }
        }

        public string SelectNzpUl(string ul)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "SELECT * FROM bill01_data.s_ulica where ulica LIKE upper('"+ul+"%') AND nzp_raj = 311005102";
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
        }

        public DataTable SelectKvar(string database, int nzp_dom)
        {
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = @"SELECT nzp_kvar FROM bill01_data.kvar where nzp_dom = " + nzp_dom;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                conn.Open();
                da.Fill(dt);
                return dt;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        public DataTable SelectChargeForEzhkh(int month, int year)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = @"SELECT ulica || ', д.' || ndom as address, nkvar, pkod, sup.name_supp, serv.service, c.tarif, '"+year+"-"+month+@"-01'::date, c.c_calc, c.sum_nedop, c.sum_charge, c.reval, c.sum_insaldo, coalesce(p.sum_rcl, 0) as sum_rcl, 
                            c.sum_money, c.sum_charge + c.reval - c.sum_nedop + coalesce(p.sum_rcl, 0) as charge, c.sum_outsaldo
                            FROM bill01_charge_"+(year - 2000).ToString("00")+@".charge_"+month.ToString("00")+@" c
                            INNER JOIN fbill_data.kvar k on k.nzp_kvar = c.nzp_kvar
                            INNER JOIN fbill_data.dom d on d.nzp_dom = k.nzp_dom 
                            INNER JOIN fbill_data.s_ulica u on u.nzp_ul = d.nzp_ul 
                            INNER JOIN fbill_kernel.supplier sup on sup.nzp_supp = c.nzp_supp
                            INNER JOIN fbill_kernel.services serv on serv.nzp_serv = c.nzp_serv
                            LEFT JOIN (SELECT nzp_kvar, nzp_serv, sum(sum_rcl) as sum_rcl from bill01_charge_"+(year - 2000).ToString("00")+@".perekidka 
	                                                where month_ = "+month+@" group by 1,2) p on p.nzp_kvar = c.nzp_kvar and p.nzp_serv = c.nzp_serv
                            where c.nzp_serv != 1 and c.dat_charge is null and (c.tarif > 0 or c.c_calc > 0 or c.sum_nedop > 0 or c.sum_charge > 0 or c.reval > 0 or c.sum_insaldo > 0 or c.sum_money > 0 
                            or c.sum_money > 0 or coalesce(p.sum_rcl, 0) > 0 or c.sum_charge + c.reval - c.sum_nedop + coalesce(p.sum_rcl, 0) > 0 or c.sum_outsaldo > 0) and k.nzp_dom = 30883
                            order by 1,2";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                conn.Open();
                da.Fill(dt);
                return dt;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        public DataTable SelectCounters(string database)
        {
            //string connStr = "Server=localhost;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = @"SELECT ul.ulica || ', д.' || d.ndom || ' кв.'|| k.nkvar, k.fio, " + (database == "billAuk" ? "k.num_ls" : "k.pkod") +
                                @", s.service, c.nzp_counter, c.num_cnt,
                                CASE WHEN Extract(month from dat_uchet) - 1 = 0 THEN 12 ELSE Extract(month from dat_uchet) - 1 END as month,
                                c.val_cnt,
                                CASE WHEN Extract(month from dat_uchet) - 1 = 0 THEN Extract(year from dat_uchet) - 1 ELSE Extract(year from dat_uchet) END as year 
                                FROM bill01_data.kvar k
                                INNER JOIN bill01_data.dom d on d.nzp_dom = k.nzp_dom
                                INNER JOIN bill01_data.s_ulica ul on d.nzp_ul = ul.nzp_ul
                                INNER JOIN bill01_data.counters c on c.nzp_kvar = k.nzp_kvar
                                INNER JOIN bill01_data.counters_spis cs on cs.nzp_counter = c.nzp_counter
                                INNER JOIN bill01_kernel.services s on s.nzp_serv = c.nzp_serv
                                where ((Extract(month from dat_uchet) >= 10 AND Extract(year from dat_uchet) = 2015) OR Extract(year from dat_uchet) >= 2016)
                                AND (cs.dat_close is null OR cs.dat_close >= current_date) and c.is_actual != 100
                                order by 1,2,4,5,9,7";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                da.Fill(dt);
                return dt;
            }
            catch (Exception e)
            {
                return null;
            }
        }

        public string SelectDom(string database, string ulica, string ndom)
        {
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = @"SELECT nzp_dom
                                FROM bill01_data.dom d
                                INNER JOIN bill01_data.s_ulica ul on ul.nzp_ul = d.nzp_ul
                                where ul.ulica = '" + ulica + "' AND ndom = '" + ndom + "'";
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
                    return "0" + "|Найдено больше или меньше 1-го дома";
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
        }

        public Dictionary<string, string> SelectPrms(string database, string prm_name)
        {
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "SELECT nzp_prm, prm_num FROM fbill_kernel.prm_name where name_prm = '" + prm_name + "'";
            Dictionary<string, string> prms = new Dictionary<string, string>();
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                conn.Open();
                da.Fill(dt);
                if (dt.Rows.Count == 0)
                {
                    return null;
                }
                else
                {
                    prms.Add("nzp_prm", dt.Rows[0][0].ToString());
                    prms.Add("prm_num", dt.Rows[0][1].ToString());
                    return prms;
                }

            }
            catch (Exception e)
            {
                return null;
            }
        }

        public void UpdateParams(string database, string nzp_prm, string prm_num, string nzp, string dateS, string val_prm)
        {
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "UPDATE bill01_data.prm_" + prm_num + " SET is_actual = 100, dat_del = CURRENT_DATE, user_del = 1 where nzp_prm = " + nzp_prm + " AND nzp = " + nzp + " AND is_actual = 1";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            cmdText = "INSERT INTO bill01_data.prm_" + prm_num + "(nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user, dat_when,  month_calc) " +
                "values(" + nzp + ", " + nzp_prm + ", to_date('" + dateS + "','dd.mm.yyyy'),  to_date('01.01.3000','dd.mm.yyyy'), '" + val_prm + "', 1, 1, CURRENT_DATE, to_date('01.01.3000','dd.mm.yyyy'))";
            cmd = new NpgsqlCommand(cmdText, conn);
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        public int InsertDom(int nzp_ul, string dom)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "INSERT INTO fbill_data.dom(nzp_land, nzp_stat, nzp_town, nzp_raj, nzp_ul, nzp_area, nzp_geu, idom, ndom, nkor, nzp_wp, pref) " +
                "values(1, 104259, 310001035, -1, " + nzp_ul + ", 2, 2, '" + dom.Replace('А', ' ').Trim() + "', '" + dom + "', '-', 25, 'bill01') returning nzp_dom";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            int nzp_dom = Convert.ToInt32(cmd.ExecuteScalar());
            cmdText = "INSERT INTO bill01_data.dom(nzp_dom, nzp_land, nzp_stat, nzp_town, nzp_raj, nzp_ul, nzp_area, nzp_geu, idom, ndom, nkor) " +
                "values(" + nzp_dom + ", 1, 104259, 310001035, -1, " + nzp_ul + ", 2, 2, '" + dom.Replace('А', ' ').Trim() + "', '" + dom + "', '-')";
            cmd = new NpgsqlCommand(cmdText, conn);
            cmd.ExecuteNonQuery();
            conn.Close();
            return nzp_dom;
        }

        public int InsertKvar(int nzp_dom, string nkvar, int num_ls, string fio, int ikvar)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "INSERT INTO fbill_data.kvar(nzp_area, nzp_geu, nzp_dom, nkvar, nkvar_n, num_ls, fio, ikvar, typek, pref, is_open, nzp_wp, area_code) " +
                "values(2, 2, " + nzp_dom + ", '" + nkvar + "', '-', " + num_ls + ", '" + fio + "', " + ikvar + ", 1, 'bill01', 1, 25, 40124) returning nzp_kvar";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            int nzp_kvar = Convert.ToInt32(cmd.ExecuteScalar());
            cmdText = "INSERT INTO bill01_data.kvar(nzp_kvar, nzp_area, nzp_geu, nzp_dom, nkvar, nkvar_n, num_ls, fio, ikvar, typek) " +
                "values(" + nzp_kvar + ", 2, 2, " + nzp_dom + ", '" + nkvar + "', '-', " + num_ls + ", '" + fio + "', " + ikvar + ", 1) returning nzp_kvar";
            cmd = new NpgsqlCommand(cmdText, conn);
            cmd.ExecuteNonQuery();
            conn.Close();
            return nzp_kvar;
        }

        public void InsertDateOpen(int nzp, string dat_s)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "INSERT INTO bill01_data.prm_3(nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user, dat_when, cur_unl, nzp_wp, month_calc) " +
                "values(" + nzp + ", 51,  to_date('" + dat_s + "','dd.mm.yyyy'), to_date('01.01.3000','dd.mm.yyyy'), '1', 1, 88888889, CURRENT_DATE, 0, 25, to_date('01.11.2014','dd.mm.yyyy'))";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        public void InsertPrm1(int nzp, string val_prm, int nzp_prm)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "INSERT INTO bill01_data.prm_1(nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user, dat_when, month_calc) " +
                "values(" + nzp + ", " + nzp_prm + ",  to_date('01.10.2014','dd.mm.yyyy'), to_date('01.01.3000','dd.mm.yyyy'), '" + val_prm + "', 1, 88888889, CURRENT_DATE, to_date('01.11.2014','dd.mm.yyyy'))";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        public int DelCounter(string ulica, string ndom, string nkvar, bool hvs, bool gvs)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            //Добавляем строку в pref_kernel.prm_name
            string cmdText = @"SELECT k.*
                            FROM bill01_data.kvar k
                            INNER JOIN bill01_data.dom d on d.nzp_dom = k.nzp_dom
                            INNER JOIN bill01_data.s_ulica u on u.nzp_ul = d.nzp_ul
                            WHERE u.ulica = '"+ulica+"' AND d.ndom = '"+ndom+"' AND k.nkvar = '"+nkvar+"'";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            int nzp_kvar = 0;
            try
            {
                da.Fill(dt);
                nzp_kvar = Convert.ToInt32(dt.Rows[0][0].ToString());
            }
            catch (Exception e)
            {
                
            }
            if (nzp_kvar != 0)
            {
                conn.Open();
                if (hvs)
                {
                    cmdText = "DELETE FROM bill01_data.counters_spis where nzp = " + nzp_kvar + " AND nzp_serv = 6";
                    cmd = new NpgsqlCommand(cmdText, conn);
                    cmd.ExecuteNonQuery();
                }
                if (gvs)
                {
                    cmdText = "DELETE FROM bill01_data.counters_spis where nzp = " + nzp_kvar + " AND nzp_serv = 9";
                    cmd = new NpgsqlCommand(cmdText, conn);
                    cmd.ExecuteNonQuery();
                }
                conn.Close();
                return 1;
            }
            else
            {
                return 0;
            }
        }

        public string SelectNzpKvar(string num_ls, string ndom, string nkvar)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "SELECT nzp_kvar FROM bill01_data.kvar k INNER JOIN bill01_data.dom d on d.nzp_dom = k.nzp_dom WHERE num_ls = " + num_ls + " AND nkvar LIKE '%"
                + nkvar + "%' AND ndom = '" + ndom + "'";
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

        public string SelectNzpKvar(string pkod)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "SELECT nzp_kvar FROM bill01_data.kvar k WHERE pkod = " + pkod;
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

        public string SelectNzpKvar(string database, string ulica, string ndom, string nkvar)
        {
            string connStr = "Server=192.168.1.25;Database="+ database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "SELECT nzp_kvar FROM bill01_data.kvar k " +
                             " INNER JOIN bill01_data.dom d on d.nzp_dom = k.nzp_dom " +
                             " INNER JOIN bill01_data.s_ulica ul on ul.nzp_ul = d.nzp_ul " +
                             " WHERE ul.ulica = '"+ ulica + "' and ndom = '"+ ndom + "' and nkvar = '"+ nkvar + "'";
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
            string cmdText = "DELETE FROM bill01_data.gilec where nzp_gil in (SELECT nzp_gil FROM bill01_data.kart where nzp_kvar = "+ nzp_kvar + ")";
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

        public List<string> SelectNzpKvarByPkod10NzpDom(string database, string pkod10, int nzp_dom, string bank)
        {
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "SELECT nzp_kvar, num_ls FROM " + bank + "_data.kvar k WHERE pkod10 = '" + pkod10 + "' AND nzp_dom = " + nzp_dom;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            List<string> data = new List<string>();
            DataTable dt = new DataTable();
            try
            {
                conn.Open();
                da.Fill(dt);
                if (dt.Rows.Count != 1)
                {
                    return null;
                }
                else
                {
                    data.Add(dt.Rows[0][0].ToString());
                    data.Add(dt.Rows[0][1].ToString());
                    return data;
                }
            }

            catch (Exception e)
            {
                return null;
            }
            finally
            {
                conn.Close();
            }
        }

        public List<string> SelectNzpKvarByPkod10NzpDomNKvar(string database, string pkod10, int nzp_dom, string bank, string nkvar)
        {
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "SELECT nzp_kvar, num_ls FROM " + bank + "_data.kvar k WHERE pkod10 = '" + pkod10 + "' AND nzp_dom = " + nzp_dom + " AND nkvar = '" + nkvar + "'";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            List<string> data = new List<string>();
            DataTable dt = new DataTable();
            try
            {
                conn.Open();
                da.Fill(dt);
                if (dt.Rows.Count != 1)
                {
                    return null;
                }
                else
                {
                    data.Add(dt.Rows[0][0].ToString());
                    data.Add(dt.Rows[0][1].ToString());
                    return data;
                }
            }

            catch (Exception e)
            {
                return null;
            }
            finally
            {
                conn.Close();
            }
        }

        public List<string> SelectNzpKvarPkod(string pkod)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "SELECT nzp_kvar, num_ls FROM bill01_data.kvar k WHERE pkod = " + pkod;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            List<string> data = new List<string>();
            try
            {
                conn.Open();
                da.Fill(dt);
                if (dt.Rows.Count != 1)
                {
                    return null;
                }
                else
                {
                    data.Add(dt.Rows[0][0].ToString());
                    data.Add(dt.Rows[0][1].ToString());
                    return data;
                }
            }

            catch (Exception e)
            {
                return null;
            }
            finally
            {
                conn.Close();
            }
        }

        public void UpdateLs()
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = @"SELECT nzp_kvar, count(nzp_kvar), max(num_ls) FROM (SELECT num_ls, nzp_kvar 
from bill01_charge_15.charge_02 
group by 1,2
order by 2,1) t1
group by 1
having count(nzp_kvar)>1";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                da.Fill(dt);
                for (int i = 0; i<dt.Rows.Count;i++)
                {
                    cmdText = "UPDATE bill01_charge_15.charge_02 SET num_ls = " + dt.Rows[i][2] + " where nzp_kvar = " + dt.Rows[i][0];
                    NpgsqlConnection conn2 = new NpgsqlConnection(connStr);
                    NpgsqlCommand cmd2 = new NpgsqlCommand(cmdText, conn2);
                    //Получаем nzp_prm добавленной нами записи
                    conn2.Open();
                    cmd2.ExecuteNonQuery();
                    conn2.Close();
                }
            }

            catch (Exception e)
            {
                string str = e.ToString();
            }
            finally
            {
                conn.Close();
            }
        }

        public string SelectNzpKvar2(string num_ls)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "SELECT nzp_kvar FROM bill01_data.kvar k WHERE num_ls = " + num_ls;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
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
        }

        public List<string> SelectNzpKvar(string database, string num_ls, string ndom, string nkvar, int t)
        {
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            //string connStr = "Server=localhost;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "SELECT nzp_kvar, num_ls FROM bill01_data.kvar k INNER JOIN bill01_data.dom d on d.nzp_dom = k.nzp_dom WHERE num_ls = " + num_ls + " AND nkvar = '"
                + nkvar + "' AND ndom = '" + ndom + "'";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            List<string> data = new List<string>();
            try
            {
                conn.Open();
                da.Fill(dt);
                if (dt.Rows.Count != 1)
                {
                    return null;
                }
                else
                {
                    data.Add(dt.Rows[0][0].ToString());
                    data.Add(dt.Rows[0][1].ToString());
                    return data;
                }
            }

            catch (Exception e)
            {
                return null;
            }
            finally
            {
                conn.Close();
            }
        }

        public List<string> SelectNzpKvar2(string database, string num_ls, string ndom, string nkvar, int t)
        {
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "SELECT nzp_kvar, num_ls FROM bill01_data.kvar k INNER JOIN bill01_data.dom d on d.nzp_dom = k.nzp_dom WHERE pkod10 = " + num_ls + " AND nkvar = '"
                + nkvar + "' AND ndom = '" + ndom + "'";
            //Console.WriteLine(cmdText);
            //Console.ReadLine();
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            List<string> data = new List<string>();
            try
            {
                conn.Open();
                da.Fill(dt);
                if (dt.Rows.Count != 1)
                {
                    return null;
                }
                else
                {
                    data.Add(dt.Rows[0][0].ToString());
                    data.Add(dt.Rows[0][1].ToString());
                    return data;
                }
            }

            catch (Exception e)
            {
                return null;
            }
            finally
            {
                conn.Close();
            }
        }

        public List<string> SelectNzpKvarByPkod(string pkod)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "SELECT nzp_kvar, num_ls FROM bill01_data.kvar WHERE pkod = " + pkod;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            List<string> data = new List<string>();
            try
            {
                conn.Open();
                da.Fill(dt);
                if (dt.Rows.Count != 1)
                {
                    return null;
                }
                else
                {
                    data.Add(dt.Rows[0][0].ToString());
                    data.Add(dt.Rows[0][1].ToString());
                    return data;
                }
            }

            catch (Exception e)
            {
                return null;
            }
            finally
            {
                conn.Close();
            }
        }

        public List<string> SelectNzpKvarByNumLs(string database, string num_ls)
        {
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "SELECT nzp_kvar, num_ls FROM bill01_data.kvar WHERE num_ls = " + num_ls;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            List<string> data = new List<string>();
            try
            {
                conn.Open();
                da.Fill(dt);
                if (dt.Rows.Count != 1)
                {
                    return null;
                }
                else
                {
                    data.Add(dt.Rows[0][0].ToString());
                    data.Add(dt.Rows[0][1].ToString());
                    return data;
                }
            }

            catch (Exception e)
            {
                return null;
            }
            finally
            {
                conn.Close();
            }
        }

        public List<string> SelectPkodByNumLs(string database, string num_ls)
        {
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "SELECT pkod FROM bill01_data.kvar WHERE num_ls = " + num_ls;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            List<string> data = new List<string>();
            try
            {
                conn.Open();
                da.Fill(dt);
                if (dt.Rows.Count != 1)
                {
                    return null;
                }
                else
                {
                    data.Add(dt.Rows[0][0].ToString());
                    return data;
                }
            }

            catch (Exception e)
            {
                return null;
            }
            finally
            {
                conn.Close();
            }
        }

       

        public int InsertGil(string database)
        {
            string connStr = "Server=192.168.1.25;Database="+ database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "INSERT INTO bill01_data.gilec(sogl) VALUES(0) returning nzp_gil";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            //Получаем nzp_prm добавленной нами записи
            int nzp_gil = Convert.ToInt32(cmd.ExecuteScalar());
            conn.Close();
            return nzp_gil;
        }

        public int InsertDocBase(string database, string comment)
        {
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            //string connStr = "Server=localhost;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "INSERT INTO fbill_data.document_base(num_doc, dat_doc, nzp_type_doc, comment) VALUES(1, to_date('2015-09-01','yyyy-mm-dd'), 4, '" + comment + "') returning nzp_doc_base";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            //Получаем nzp_prm добавленной нами записи
            int nzp_doc_base = Convert.ToInt32(cmd.ExecuteScalar());
            conn.Close();
            return nzp_doc_base;
        }

        public int InsertKart(string database, int nzp_gil, string nzp_kvar, string fam, string ima, string otch, string dat_rog, string gender, int nzp_dok, string serij, string nomer, 
            string vid_dat, string vid_mes, string tprp, string dat_sost, string dat_reg, int nzp_rod, string rodstvo,
            string region_op, string okrug_op, string gorod_op, string npunkt_op, string region_ku, string okrug_ku, string gorod_ku, string rem_ku)
        {
            string connStr = "Server=192.168.1.25;Database="+ database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
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
                "VALUES("+nzp_gil+", 1, '"+fam+"', '"+ima+"', '"+otch+"', to_date('"+dat_rog+"', 'dd.mm.yyyy'), '"+gender+"', '"+tprp+"', 1, 1, "+ nzp_tkrt + ", "+nzp_kvar+", -1, "+nzp_rod
                +", "+nzp_dok+", '"+serij+"', '"+nomer+"', '"+vid_mes+"', to_date('"+vid_dat+"', 'dd.mm.yyyy'),  to_date('"+ dat_sost + "', 'dd.mm.yyyy'),  to_date('"+dat_reg
                +"', 'dd.mm.yyyy'), current_date, 1, 1, '"+rodstvo+"','"+ region_op + "', '"+ okrug_op + "', '"+ gorod_op + "', '"+ npunkt_op + "', '"+ region_ku + "', '" 
                + okrug_ku + "', '"+ gorod_ku + "', '"+ rem_ku + "') returning nzp_kart";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            //Получаем nzp_prm добавленной нами записи
            int nzp_kart = Convert.ToInt32(cmd.ExecuteScalar());
            conn.Close();
            return nzp_kart;
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

        public void InsertPerekidka(string database, int nzp_kvar, decimal sum_rcl, int nzp_doc_base, int num_ls, int nzp_serv, int nzp_supp, string date_rcl, int month_, int year_)
        {
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            //string connStr = "Server=localhost;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "INSERT INTO bill01_charge_" + (year_ - 2000).ToString("00") + ".perekidka(nzp_kvar, num_ls, nzp_serv, nzp_supp, type_rcl, date_rcl, tarif, volum, sum_rcl, month_, nzp_user, nzp_doc_base) " +
                "VALUES(" + nzp_kvar + ", " + num_ls + ", " + nzp_serv + ", " + nzp_supp + ", 102, to_date('" + date_rcl + "','yyyy-mm-dd'), 0, 0, " + sum_rcl.ToString().Replace(',', '.') + ", " + month_ + ", 1, " + nzp_doc_base + ")";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        public void InsertPerekidka2(int nzp_kvar, decimal sum_rcl, int nzp_doc_base, int num_ls)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "INSERT INTO bill01_charge_15.perekidka(nzp_kvar, num_ls, nzp_serv, nzp_supp, type_rcl, date_rcl, tarif, volum, sum_rcl, month_, nzp_user, nzp_doc_base) " +
                "VALUES(" + nzp_kvar + ", " + num_ls + ", 9, 101178, 102, to_date('2015-06-25','yyyy-mm-dd'), 0, 0, " + sum_rcl + ", 6, 88888896, " + nzp_doc_base + ")";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        public void InsertPerekidka3(int nzp_kvar, decimal sum_rcl, int nzp_doc_base, int num_ls)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "INSERT INTO bill01_charge_15.perekidka(nzp_kvar, num_ls, nzp_serv, nzp_supp, type_rcl, date_rcl, tarif, volum, sum_rcl, month_, nzp_user, nzp_doc_base) " +
                "VALUES(" + nzp_kvar + ", " + num_ls + ", 8, 101185, 102, to_date('2015-05-28','yyyy-mm-dd'), 0, 0, " + sum_rcl + ", 6, 88888896, " + nzp_doc_base + ")";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        public void InsertPerekidka4(int nzp_kvar, decimal sum_rcl, int nzp_doc_base, int num_ls)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "INSERT INTO bill01_charge_15.perekidka(nzp_kvar, num_ls, nzp_serv, nzp_supp, type_rcl, date_rcl, tarif, volum, sum_rcl, month_, nzp_user, nzp_doc_base) " +
                "VALUES(" + nzp_kvar + ", " + num_ls + ", 9, 101184, 102, to_date('2015-06-25','yyyy-mm-dd'), 0, 0, " + sum_rcl + ", 6, 88888896, " + nzp_doc_base + ")";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        public void InsertPerekidka5(int nzp_kvar, decimal sum_rcl, int nzp_doc_base, int num_ls)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "INSERT INTO bill01_charge_15.perekidka(nzp_kvar, num_ls, nzp_serv, nzp_supp, type_rcl, date_rcl, tarif, volum, sum_rcl, month_, nzp_user, nzp_doc_base) " +
                "VALUES(" + nzp_kvar + ", " + num_ls + ", 25, 101184, 102, to_date('2015-06-30','yyyy-mm-dd'), 0, 0, " + sum_rcl + ", 7, 88888896, " + nzp_doc_base + ")";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        public void InsertPerekidka6(int nzp_kvar, decimal sum_rcl, int nzp_doc_base, int num_ls)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "INSERT INTO bill01_charge_15.perekidka(nzp_kvar, num_ls, nzp_serv, nzp_supp, type_rcl, date_rcl, tarif, volum, sum_rcl, month_, nzp_user, nzp_doc_base) " +
                "VALUES(" + nzp_kvar + ", " + num_ls + ", 500, 101185, 102, to_date('2015-07-14','yyyy-mm-dd'), 0, 0, " + sum_rcl + ", 7, 88888896, " + nzp_doc_base + ")";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        public void InsertPerekidka7(int nzp_kvar, decimal sum_rcl, int nzp_doc_base, int num_ls)
        {
            string connStr = "Server=localhost;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "INSERT INTO bill01_charge_15.perekidka(nzp_kvar, num_ls, nzp_serv, nzp_supp, type_rcl, date_rcl, tarif, volum, sum_rcl, month_, nzp_user, nzp_doc_base) " +
                "VALUES(" + nzp_kvar + ", " + num_ls + ", 100015, 101179, 102, to_date('2015-09-25','yyyy-mm-dd'), 0, 0, " + sum_rcl.ToString().Replace(',','.') + ", 9, 88888896, " + nzp_doc_base + ")";
            Console.WriteLine(cmdText);
            Console.ReadLine();
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        public void InsertPerekidka14(int nzp_kvar, decimal sum_rcl, int nzp_doc_base, int num_ls)
        {
            string connStr = "Server=localhost;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "INSERT INTO bill01_charge_15.perekidka(nzp_kvar, num_ls, nzp_serv, nzp_supp, type_rcl, date_rcl, tarif, volum, sum_rcl, month_, nzp_user, nzp_doc_base) " +
                "VALUES(" + nzp_kvar + ", " + num_ls + ", 9, 101185, 102, to_date('2015-09-25','yyyy-mm-dd'), 0, 0, " + sum_rcl.ToString().Replace(',','.') + ", 9, 88888896, " + nzp_doc_base + ")";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        public void InsertPerekidkaByNzpServAndMonthAndSupp(string database, int nzp_kvar, decimal sum_rcl, int nzp_doc_base, int num_ls, string nzp_serv, int month, string nzp_supp, int year = 2015, string bank = "bill01")
        {
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "INSERT INTO " + bank + "_charge_" + (year - 2000).ToString("00") + ".perekidka(nzp_kvar, num_ls, nzp_serv, nzp_supp, type_rcl, date_rcl, tarif, volum, sum_rcl, month_, nzp_user, nzp_doc_base) " +
                "VALUES(" + nzp_kvar + ", " + num_ls + ", " + nzp_serv + ", " + nzp_supp + ", 102, to_date('2016-01-25','yyyy-mm-dd'), 0, 0, " + sum_rcl.ToString().Replace(',', '.') + ", " + 
                month + ", 1, " + nzp_doc_base + ")";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        public void InsertFio(string nkvar, string num_ls, string fio)
        {
            string connStr = "Server=192.168.1.25;Database=billDemidov;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "INSERT INTO fbill_data.kvar(nzp_area, nzp_geu, nzp_dom, nkvar, nkvar_n, num_ls, fio, ikvar, typek, pref, is_open, nzp_wp, area_code) " +
                "values(2, 2, 7155109, '" + nkvar + "', '-', " + num_ls + ", '" + fio + "', " + nkvar + ", 1, 'bill01', 1, 25, 40124) returning nzp_kvar";
            Console.WriteLine(cmdText);
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            int nzp_kvar = Convert.ToInt32(cmd.ExecuteScalar());
            cmdText = "INSERT INTO bill01_data.kvar(nzp_kvar, nzp_area, nzp_geu, nzp_dom, nkvar, nkvar_n, num_ls, fio, ikvar, typek) " +
                "values(" + nzp_kvar + ", 2, 2, 7155109, '" + nkvar + "', '-', " + num_ls + ", '" + fio + "', " + nkvar + ", 1) returning nzp_kvar";
            Console.WriteLine(cmdText);
            cmd = new NpgsqlCommand(cmdText, conn);
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        public void InsertFio(string nkvar, string num_ls, string fio, string nzp_dom)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "INSERT INTO fbill_data.kvar(nzp_area, nzp_geu, nzp_dom, nkvar, nkvar_n, fio, ikvar, typek, pref, is_open, nzp_wp, area_code) " +
                "values(2, 2, "+ nzp_dom + ", '" + nkvar + "', '-', '" + fio + "', " + nkvar + ", 1, 'bill01', 1, 25, 40124) returning nzp_kvar";
            //Console.WriteLine(cmdText);
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            int nzp_kvar = Convert.ToInt32(cmd.ExecuteScalar());
            cmdText = "UPDATE fbill_data.kvar set num_ls = "+ nzp_kvar + " where nzp_kvar = " + nzp_kvar;
            cmd = new NpgsqlCommand(cmdText, conn);
            cmd.ExecuteNonQuery();

            cmdText = "INSERT INTO bill01_data.kvar(nzp_kvar, nzp_area, nzp_geu, nzp_dom, nkvar, nkvar_n, fio, ikvar, typek) " +
                "values(" + nzp_kvar + ", 2, 2, " + nzp_dom + ", '" + nkvar + "', '-', '" + fio + "', " + nkvar + ", 1) returning nzp_kvar";
           // Console.WriteLine(cmdText);
            cmd = new NpgsqlCommand(cmdText, conn);
            cmd.ExecuteNonQuery();
            cmdText = "UPDATE bill01_data.kvar set num_ls = " + nzp_kvar + " where nzp_kvar = " + nzp_kvar;
            cmd = new NpgsqlCommand(cmdText, conn);
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        public Int32 InsertFio(string nkvar, string fio)
        {
            //string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string connStr = "Server=localhost;Database=billTlt;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "INSERT INTO fbill_data.kvar(nzp_area, nzp_geu, nzp_dom, nkvar, nkvar_n, num_ls, fio, ikvar, typek, pref, is_open, nzp_wp, area_code) " +
                             "VALUES (2, 2, 7155105, '" + nkvar + "', '-', 0, '" + fio + "', " + nkvar + ", 1, 'bill02', 1, 25, 40102) returning nzp_kvar";
            Console.WriteLine(cmdText);

            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            int nzp_kvar = Convert.ToInt32(cmd.ExecuteScalar());

            cmdText = "UPDATE fbill_data.kvar set num_ls = " + nzp_kvar + " where nzp_kvar = " + nzp_kvar;
            cmd = new NpgsqlCommand(cmdText, conn);
            cmd.ExecuteNonQuery();

            cmdText = @" INSERT INTO bill02_data.kvar(nzp_kvar, nzp_area, nzp_geu, nzp_dom, nkvar, nkvar_n, num_ls, porch, phone, dat_notp_s, dat_notp_po, fio, ikvar, uch, gil_s, 
                        remark, typek, pkod, pkod10) 
                        SELECT nzp_kvar, nzp_area, nzp_geu, nzp_dom, nkvar, nkvar_n, num_ls, porch, phone, dat_notp_s, dat_notp_po, fio, ikvar, uch, gil_s, remark, typek, pkod, pkod10 
                        FROM fbill_data.kvar where nzp_kvar = " + nzp_kvar;
            //Console.WriteLine(cmdText);
            cmd = new NpgsqlCommand(cmdText, conn);
            cmd.ExecuteNonQuery();
            conn.Close();
            return nzp_kvar;
        }

        public void UpdateFio(string database, string nzp_kvar, string fio)
        {
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "UPDATE fbill_data.kvar set fio = '" + fio + "' where nzp_kvar = " + nzp_kvar;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            cmdText = "UPDATE bill01_data.kvar set fio = '" + fio + "' where nzp_kvar = " + nzp_kvar;
            cmd = new NpgsqlCommand(cmdText, conn);
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        public DataTable TestFio(string database, string nzp_kvar)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "SELECT * FROM fbill_data.kvar where nzp_kvar = " + nzp_kvar;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }

        public void InsertPrm(string nkvar, string prm4, string prm5, string prm1010270)
        {
            string connStr = "Server=192.168.1.25;Database=billDemidov;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "SELECT nzp_kvar FROM bill01_data.kvar where nzp_dom = 7155108 AND nkvar = '" + nkvar + "'";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            int nzp_kvar = Convert.ToInt32(dt.Rows[0][0]);
            cmdText = "INSERT INTO bill01_data.prm_1(nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user, dat_when) " +
                "values(" + nzp_kvar + ", 4, to_date('01.04.2015','dd.mm.yyyy'), to_date('01.01.3000','dd.mm.yyyy'), '" + prm4 + "', 1, 133, to_date('27.07.2015','dd.mm.yyyy'))";
            cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            cmdText = "INSERT INTO bill01_data.prm_1(nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user, dat_when) " +
               "values(" + nzp_kvar + ", 5, to_date('01.04.2015','dd.mm.yyyy'), to_date('01.01.3000','dd.mm.yyyy'), '" + prm5 + "', 1, 133, to_date('27.07.2015','dd.mm.yyyy'))";
            cmd = new NpgsqlCommand(cmdText, conn);
            cmd.ExecuteNonQuery();
            cmdText = "INSERT INTO bill01_data.prm_1(nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user, dat_when) " +
               "values(" + nzp_kvar + ", 1010270, to_date('01.04.2015','dd.mm.yyyy'), to_date('01.01.3000','dd.mm.yyyy'), '" + prm1010270 + "', 1, 133, to_date('27.07.2015','dd.mm.yyyy'))";
            cmd = new NpgsqlCommand(cmdText, conn);
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        public int InsertPrm(Int32 num_ls, string val_prm, Int32 nzp_prm)
        {
            string connStr = "Server=192.168.1.25;Database=billTlt;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "SELECT nzp_kvar FROM bill01_data.kvar where num_ls = " + num_ls;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            if (dt.Rows.Count != 1)
                return 0;
            int nzp_kvar = Convert.ToInt32(dt.Rows[0][0]);
            cmdText = "INSERT INTO bill01_data.prm_1(nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user, dat_when) " +
                "values(" + nzp_kvar + ", " + nzp_prm + ", to_date('01.09.2015','dd.mm.yyyy'), to_date('01.01.3000','dd.mm.yyyy'), '" + val_prm + "', 1, 1, to_date('01.09.2015','dd.mm.yyyy'))";
            cmd = new NpgsqlCommand(cmdText, conn);
            try
            {
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
                return 1;
            }
            catch (Exception)
            {
                return 0;
            }
           
        }

        public int InsertPrm(string database, string nzp_kvar, string val_prm, Int32 nzp_prm)
        {
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = cmdText = "INSERT INTO bill01_data.prm_1(nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user, dat_when) " +
                "values(" + nzp_kvar + ", " + nzp_prm + ", to_date('01.10.2015','dd.mm.yyyy'), to_date('01.01.3000','dd.mm.yyyy'), '" + val_prm + "', 1, 1, to_date('01.10.2015','dd.mm.yyyy'))";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            try
            {
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
                return 1;
            }
            catch (Exception)
            {
                return 0;
            }

        }

        public int InsertPrmByKvar(Int32 nzp, string val_prm, Int32 nzp_prm)
        {
            //string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string connStr = "Server=localhost;Database=billTlt;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "INSERT INTO bill02_data.prm_1(nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user, dat_when) " +
                "values(" + nzp + ", " + nzp_prm + ", to_date('01.01.2015','dd.mm.yyyy'), to_date('01.01.3000','dd.mm.yyyy'), '" + val_prm + "', 1, 1, to_date('01.01.2015','dd.mm.yyyy'))";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            try
            {
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
                return 1;
            }
            catch (Exception)
            {
                return 0;
            }

        }

        public List<string> SelectKvarParams(String localhost, String database, String nkvar, Int32 nzp_dom)
        {
            //string connStr = "Server=localhost;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string connStr = "Server=" + (localhost == "localhost" ? "localhost" : "192.168.1.25") + ";Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "SELECT nzp_kvar, num_ls FROM fbill_data.kvar where is_open = '1' and nkvar = '" + nkvar + "' AND nzp_dom = " + nzp_dom;
            List<string> kvarParams = new List<string>();
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                da.Fill(dt);
                if (dt.Rows.Count == 0)
                    return null;
                else
                {
                    kvarParams.Add(dt.Rows[0][0].ToString());
                    kvarParams.Add(dt.Rows[0][1].ToString());
                    return kvarParams;
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        public List<string> SelectKvarParamsByNumLs(String num_ls, Int32 nzp_dom)
        {
            string connStr = "Server=192.168.1.25;Database=billTlt;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "SELECT nzp_kvar, num_ls FROM bill01_data.kvar where num_ls = '" + num_ls + "' AND nzp_dom = " + nzp_dom;
            List<string> kvarParams = new List<string>();
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                da.Fill(dt);
                if (dt.Rows.Count == 0)
                    return null;
                else
                {
                    kvarParams.Add(dt.Rows[0][0].ToString());
                    kvarParams.Add(dt.Rows[0][1].ToString());
                    return kvarParams;
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

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

        public Int32 InsertCounter(string nzp, Int32 nzp_serv, string num_cnt, Int32 nzp_cnt, string dat_prov)
        {
            string connStr = "Server=192.168.1.25;Database=billTlt;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText;
            if (dat_prov != "")
            {
                cmdText = @"INSERT INTO bill01_data.counters_spis(
            nzp_type, nzp, nzp_serv, nzp_cnttype, num_cnt, is_gkal, 
            kod_pu, kod_info, 
            is_actual, nzp_cnt, nzp_user, dat_when,
            cnt_ls, dat_prov)
    VALUES (3, " + nzp + ", " + nzp_serv + ", 17, '" + num_cnt + @"', 0, 
            0, 0,
            1, " + nzp_cnt + @", 1, current_date, 
            0, '" + dat_prov + "') returning nzp_counter";
            }
            else
            {
                cmdText = @"INSERT INTO bill01_data.counters_spis(
            nzp_type, nzp, nzp_serv, nzp_cnttype, num_cnt, is_gkal, 
            kod_pu, kod_info, 
            is_actual, nzp_cnt, nzp_user, dat_when,
            cnt_ls)
    VALUES (3, " + nzp + ", " + nzp_serv + ", 17, '" + num_cnt + @"', 0, 
            0, 0,
            1, " + nzp_cnt + @", 1, current_date, 
            0) returning nzp_counter";
            }
            //Console.WriteLine(cmdText);
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            int nzp_counter = Convert.ToInt32(cmd.ExecuteScalar());
            conn.Close();
            return nzp_counter;
        }

        public Int32 UpdateCounters(string nzp, Int32 nzp_serv, string num_cnt, string date_from, string val_prm)
        {
            string connStr = "Server=192.168.1.25;Database=billTlt;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = @"SELECT nzp_counter FROM bill01_data.counters_spis where nzp = " + nzp + " and num_cnt = '" + num_cnt + "' and nzp_serv = " + nzp_serv;
            //Console.WriteLine(cmdText);
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                da.Fill(dt);
                if (dt.Rows.Count == 1)
                {
                    Int32 nzp_counter = Convert.ToInt32(dt.Rows[0][0]);
                    if (val_prm != "")
                    {
                        cmdText = @"INSERT INTO bill01_data.prm_17(nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user, dat_when)
                                   VALUES (" + nzp_counter + ", 2025, '2015-08-01', '3000-01-01', '" + val_prm.Substring(0,10) + "', 1, 1, '2015-09-20')";
                        cmd = new NpgsqlCommand(cmdText, conn);
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                    if (date_from != "")
                    {
                        cmdText = @"INSERT INTO bill01_data.counters_bounds(nzp_counter, type_id, date_from, date_to, is_actual, created_by, created_on)
                                    VALUES (" + nzp_counter + ", 2, '" + date_from.Substring(0,10) + "', '" + date_from.Substring(0,10) + "', true, 1, '2015-09-20');";
                        cmd = new NpgsqlCommand(cmdText, conn);
                        conn.Open();
                        cmd.ExecuteNonQuery();
                        conn.Close();
                    }
                    return 1;
                }
                else
                {
                    return 2;
                }
            }
            catch (Exception)
            {
                return 0;
            }
        }

        public Int32 UpdateCountersDatClose(string database, string nzp, Int32 nzp_serv, string num_cnt, string val_prm)
        {
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = @"SELECT nzp_counter FROM bill01_data.counters_spis where nzp = " + nzp + " and num_cnt = '" + num_cnt + "' and nzp_serv = " + nzp_serv;
            //Console.WriteLine(cmdText);
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                da.Fill(dt);
                if (dt.Rows.Count == 1)
                {
                    Int32 nzp_counter = Convert.ToInt32(dt.Rows[0][0]);
                    if (val_prm != "")
                    {
                        cmdText = @"SELECT nzp FROM bill01_data.prm_17 where nzp = " + nzp_counter + " and nzp_prm = 2025 and is_actual = 1";
                        cmd = new NpgsqlCommand(cmdText, conn);
                        da = new NpgsqlDataAdapter(cmd);
                        dt = new DataTable();
                        try
                        {
                            da.Fill(dt);
                            if (dt.Rows.Count >= 1)
                            {
                                cmdText = @"UPDATE bill01_data.prm_17 SET val_prm = '" + val_prm.Substring(0, 10) + "' where nzp = " + Convert.ToInt32(dt.Rows[0][0]);
                                cmd = new NpgsqlCommand(cmdText, conn);
                                conn.Open();
                                cmd.ExecuteNonQuery();
                                conn.Close();
                            }
                            else
                            {
                                cmdText = @"INSERT INTO bill01_data.prm_17(nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user, dat_when)
                                   VALUES (" + nzp_counter + ", 2025, '2015-08-01', '3000-01-01', '" + val_prm.Substring(0, 10) + "', 1, 1, '2015-09-20')";
                                cmd = new NpgsqlCommand(cmdText, conn);
                                conn.Open();
                                cmd.ExecuteNonQuery();
                                conn.Close();
                            }
                        }
                        catch (Exception)
                        {
                            return 0;
                        }
                    }
                    return 1;
                }
                else
                {
                    return 2;
                }
            }
            catch (Exception)
            {
                return 0;
            }
        }

        public Int32 InsertDatPov(Int32 nzp_counter, string date_from)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            try
            {
                if (date_from != "")
                {
                    string cmdText = @"INSERT INTO bill01_data.counters_bounds(nzp_counter, type_id, date_from, date_to, is_actual, created_by, created_on)
                                    VALUES (" + nzp_counter + ", 2, '" + date_from.Substring(0, 10) + "', '" + date_from.Substring(0, 10) + "', true, 1, '2015-09-20');";
                    NpgsqlConnection conn = new NpgsqlConnection(connStr);
                    NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
                    conn.Open();
                    cmd.ExecuteNonQuery();
                    conn.Close();
                }

                return 1;
            }
            catch (Exception)
            {
                return 0;
            }
        }

        public Int32 SelectNzpCounter(string localhost, string database, string nzp, Int32 nzp_serv, string num_cnt, string dat_pov, string old_value, string new_value = "999999999")
        {
            //string connStr = "Server=localhost;Database=billTlt;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string connStr = "Server=" + (localhost == "localhost" ? "localhost" : "192.168.1.25") + ";Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = @"SELECT s.nzp_counter,  dat_uchet
                            FROM bill01_data.counters_spis s 
                            INNER JOIN bill01_data.counters c on c.nzp_counter = s.nzp_counter  where nzp = " + nzp + " and s.nzp_serv = " + nzp_serv + " and s.num_cnt = '" + num_cnt + "' " +
                             " and (s.dat_close >= '" + dat_pov + "' or s.dat_close is null) and c.dat_uchet < to_date('" + dat_pov + "', 'yyyy-mm-dd') order by 2 DESC";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                da.Fill(dt);
                if (dt.Rows.Count == 0)
                {
                    return 0;                
                }                 
                else
                {
                    string nzp_counter = dt.Rows[0][0].ToString();
                    string dat_uchet = dt.Rows[0][1].ToString();
                    cmdText = @"SELECT s.nzp_counter
                            FROM bill01_data.counters_spis s 
                            INNER JOIN bill01_data.counters c on c.nzp_counter = s.nzp_counter  where nzp = " + nzp + " and s.nzp_serv = " + nzp_serv + " and s.num_cnt = '" + num_cnt + "' " +
                             " and (s.dat_close >= '" + dat_pov + "' or s.dat_close is null) and c.dat_uchet = to_date('" + dat_uchet + "', 'dd-mm-yyyy') and c.val_cnt <= " + new_value.Replace(',', '.');
                    //Console.WriteLine(cmdText);
                    cmd = new NpgsqlCommand(cmdText, conn);
                    da = new NpgsqlDataAdapter(cmd);
                    dt = new DataTable();
                    try
                    {
                        da.Fill(dt);
                        if (dt.Rows.Count == 0)
                        {
                            return 0;
                        }
                        else
                        {
                            return Convert.ToInt32(dt.Rows[0][0].ToString());
                        }
                    }
                    catch (Exception)
                    {
                        return 0;
                    }
                    
                }
            }
            catch (Exception)
            {
                return 0;
            }
        }

        public int InsertCounterVal(String localhost, String database, String nzp_kvar, String num_ls, Int32 nzp_serv, String num_cnt, String dat_uchet, String val_cnt, Int32 nzp_counter)
        {
            //string connStr = "Server=localhost;Database=billTlt;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string connStr = "Server=" + (localhost == "localhost" ? "localhost" : "192.168.1.25") + ";Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            Decimal val = 0;
            if (val_cnt.Trim() != "")
                val = Convert.ToDecimal(val_cnt.Replace('.',','));
            if (val == 0)
                return 1;
            string cmdText = @"INSERT INTO bill01_data.counters(nzp_kvar, num_ls, nzp_serv, nzp_cnttype, num_cnt, dat_uchet, val_cnt, is_actual, nzp_user, dat_when, cur_unl, ist, nzp_counter)
                        VALUES (" + nzp_kvar + ", " + num_ls + ", " + nzp_serv + ", 17, '" + num_cnt + "', to_date('" + dat_uchet + "', 'dd.mm.yyyy'), " + val.ToString().Replace(',','.') + ", 1, 1, current_date, 1, 0, " + nzp_counter + ")";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            cmd = new NpgsqlCommand(cmdText, conn);
            try
            {
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
                return 1;
            }
            catch (Exception)
            {
                return 0;
            }

        }

        public int InsertCounterValWithDelOld(String nzp_kvar, String num_ls, Int32 nzp_serv, String num_cnt, String dat_uchet, String val_cnt, Int32 nzp_counter)
        {
            string connStr = "Server=192.168.1.25;Database=billTlt;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            Decimal val = 0;
            if (val_cnt.Trim() != "")
                val = Convert.ToDecimal(val_cnt);
            if (val == 0)
                return 1;
            string cmdText = @"DELETE FROM bill01_data.counters WHERE nzp_kvar = " + nzp_kvar +
                             " AND num_ls = " + num_ls + " AND nzp_serv = " + nzp_serv + " AND num_cnt = '" + num_cnt +
                             "' AND dat_uchet = to_date('" + dat_uchet + "', 'dd.mm.yyyy') AND nzp_counter = " +
                             nzp_counter;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();

            cmdText = @"INSERT INTO bill01_data.counters(nzp_kvar, num_ls, nzp_serv, nzp_cnttype, num_cnt, dat_uchet, val_cnt, is_actual, nzp_user, dat_when, cur_unl, ist, nzp_counter)
                        VALUES (" + nzp_kvar + ", " + num_ls + ", " + nzp_serv + ", 17, '" + num_cnt + "', to_date('" + dat_uchet + "', 'dd.mm.yyyy'), " + val + ", 1, 1, current_date, 1, 0, " + nzp_counter + ")";
            cmd = new NpgsqlCommand(cmdText, conn);
            try
            {
                cmd.ExecuteNonQuery();
                conn.Close();
                return 1;
            }
            catch (Exception)
            {
                return 0;
            }

        }

        public int InsertOutSaldo(string nzp_kvar, string num_ls, Int32 nzp_serv, decimal sum_outsaldo)
        {
            string connStr = "Server=192.168.1.25;Database=billTlt;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            Decimal val = 0;

            string cmdText = @"SELECT * FROM bill01_charge_15.charge_08 where nzp_kvar = " + nzp_kvar + " AND num_ls = " + num_ls + " AND nzp_serv = " + nzp_serv + " AND nzp_serv != 25";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            if (dt.Rows.Count > 0)
                return 2;

            cmdText = @"INSERT INTO bill01_charge_15.charge_08(nzp_kvar, num_ls, nzp_serv, sum_outsaldo)
                                VALUES (" + nzp_kvar + ", " + num_ls + ", " + nzp_serv + ", " + sum_outsaldo + ")";
            cmd = new NpgsqlCommand(cmdText, conn);
            try
            {
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
                return 1;
            }
            catch (Exception)
            {
                return 0;
            }

        }

        public int InsertOutSaldo(string database, string nzp_kvar, string num_ls, Int32 nzp_serv, decimal sum_outsaldo, int nzp_supp)
        {
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";

            string cmdText = @"";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);

            cmdText = @"INSERT INTO bill01_charge_15.charge_08(nzp_kvar, num_ls, nzp_serv, sum_outsaldo, nzp_supp)
                                VALUES (" + nzp_kvar + ", " + num_ls + ", " + nzp_serv + ", " + sum_outsaldo.ToString().Replace(',','.') + ", " + nzp_supp + ")";
            cmd = new NpgsqlCommand(cmdText, conn);
            try
            {
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
                return 1;
            }
            catch (Exception)
            {
                return 0;
            }

        }

        public int InsertOutSaldoPeni(string database, string nzp_kvar, string num_ls, Int32 nzp_serv, decimal sum_outsaldo, int nzp_supp)
        {
            string connStr = "Server=192.168.1.25;Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            Decimal val = 0;

            string cmdText = @"SELECT * FROM bill01_charge_15.charge_08 where nzp_kvar = " + nzp_kvar + " AND num_ls = " + num_ls + " AND nzp_serv = " + nzp_serv +
                " AND nzp_serv != 25 AND nzp_serv != 515 AND nzp_serv != 2 AND nzp_serv != 7 AND nzp_serv != 8";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            if (dt.Rows.Count > 0)
                return 2;

            cmdText = @"INSERT INTO bill01_charge_15.charge_08(nzp_kvar, num_ls, nzp_serv, sum_outsaldo, nzp_supp)
                                VALUES (" + nzp_kvar + ", " + num_ls + ", " + nzp_serv + ", " + sum_outsaldo.ToString().Replace(',', '.') + ", " + nzp_supp + ")";
            cmd = new NpgsqlCommand(cmdText, conn);
            try
            {
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
                return 1;
            }
            catch (Exception)
            {
                return 0;
            }

        }

        public int DelSaldo(string nzp_kvar)
        {
            string connStr = "Server=192.168.1.25;Database=billTlt;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = @"DELETE FROM bill01_charge_15.charge_08 where nzp_kvar = " + nzp_kvar;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            try
            {
                conn.Open();
                cmd.ExecuteNonQuery();
                conn.Close();
                return 1;
            }
            catch (Exception)
            {
                return 0;
            }

        }

        public int InsertPrm(string nkvar, string fio, string prm6, string prm4, string prm8, string prm5)
        {
            string connStr = "Server=192.168.1.25;Database=billDemidov;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "SELECT nzp_kvar FROM bill01_data.kvar where nzp_dom = 7155109 AND nkvar = '" + nkvar + "' and upper(FIO) = '" +
                fio.ToUpper() + "'";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            if (dt.Rows.Count != 1)
                return 0;
            int nzp_kvar = Convert.ToInt32(dt.Rows[0][0]);
            cmdText = "INSERT INTO bill01_data.prm_1(nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user, dat_when) " +
                "values(" + nzp_kvar + ", 6, to_date('01.04.2015','dd.mm.yyyy'), to_date('01.01.3000','dd.mm.yyyy'), '" + prm6 + "', 1, 133, to_date('27.07.2015','dd.mm.yyyy'))";
            Console.WriteLine(cmdText);
            cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            cmdText = "INSERT INTO bill01_data.prm_1(nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user, dat_when) " +
               "values(" + nzp_kvar + ", 4, to_date('01.04.2015','dd.mm.yyyy'), to_date('01.01.3000','dd.mm.yyyy'), '" + prm4 + "', 1, 133, to_date('27.07.2015','dd.mm.yyyy'))";
            cmd = new NpgsqlCommand(cmdText, conn);
            cmd.ExecuteNonQuery();
            cmdText = "INSERT INTO bill01_data.prm_1(nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user, dat_when) " +
               "values(" + nzp_kvar + ", 1010270, to_date('01.04.2015','dd.mm.yyyy'), to_date('01.01.3000','dd.mm.yyyy'), '" + prm5+ "', 1, 133, to_date('27.07.2015','dd.mm.yyyy'))";
            cmd = new NpgsqlCommand(cmdText, conn);
            cmd.ExecuteNonQuery();
            cmdText = "INSERT INTO bill01_data.prm_1(nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user, dat_when) " +
               "values(" + nzp_kvar + ", 8, to_date('01.04.2015','dd.mm.yyyy'), to_date('01.01.3000','dd.mm.yyyy'), '" + prm8 + "', 1, 133, to_date('27.07.2015','dd.mm.yyyy'))";
            cmd = new NpgsqlCommand(cmdText, conn);
            cmd.ExecuteNonQuery();
            cmdText = "INSERT INTO bill01_data.prm_1(nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user, dat_when) " +
               "values(" + nzp_kvar + ", 5, to_date('01.04.2015','dd.mm.yyyy'), to_date('01.01.3000','dd.mm.yyyy'), '" + prm5 + "', 1, 133, to_date('27.07.2015','dd.mm.yyyy'))";
            cmd = new NpgsqlCommand(cmdText, conn);
            cmd.ExecuteNonQuery();
            conn.Close();
            return 1;
        }

        public int InsertPrm(string nkvar, string nzp_dom, string fio, string prm6, string prm4, string prm8, string prm5)
        {
            string connStr = "Server=192.168.1.25;Database=billAuk;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "SELECT nzp_kvar FROM bill01_data.kvar where nzp_dom = "+ nzp_dom + " AND nkvar = '" + nkvar + "' and upper(FIO) = '" +
                fio.ToUpper() + "'";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            if (dt.Rows.Count != 1)
                return 0;
            int nzp_kvar = Convert.ToInt32(dt.Rows[0][0]);
            cmdText = "INSERT INTO bill01_data.prm_1(nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user, dat_when) " +
                "values(" + nzp_kvar + ", 6, to_date('01.04.2016','dd.mm.yyyy'), to_date('01.01.3000','dd.mm.yyyy'), '" + prm6 + "', 1, 1, to_date('01.04.2016','dd.mm.yyyy'))";
            Console.WriteLine(cmdText);
            cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            cmdText = "INSERT INTO bill01_data.prm_1(nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user, dat_when) " +
               "values(" + nzp_kvar + ", 4, to_date('01.04.2016','dd.mm.yyyy'), to_date('01.01.3000','dd.mm.yyyy'), '" + prm4 + "', 1, 1, to_date('01.04.2016','dd.mm.yyyy'))";
            cmd = new NpgsqlCommand(cmdText, conn);
            cmd.ExecuteNonQuery();
            cmdText = "INSERT INTO bill01_data.prm_1(nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user, dat_when) " +
               "values(" + nzp_kvar + ", 1010270, to_date('01.04.2016','dd.mm.yyyy'), to_date('01.01.3000','dd.mm.yyyy'), '" + prm5 + "', 1, 1, to_date('01.04.2016','dd.mm.yyyy'))";
            cmd = new NpgsqlCommand(cmdText, conn);
            cmd.ExecuteNonQuery();
            cmdText = "INSERT INTO bill01_data.prm_1(nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user, dat_when) " +
               "values(" + nzp_kvar + ", 8, to_date('01.04.2016','dd.mm.yyyy'), to_date('01.01.3000','dd.mm.yyyy'), '" + 
               (prm8 == "10" ? 1 : 0) +
               "', 1, 1, to_date('01.04.2016','dd.mm.yyyy'))";
            cmd = new NpgsqlCommand(cmdText, conn);
            cmd.ExecuteNonQuery();
            cmdText = "INSERT INTO bill01_data.prm_1(nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user, dat_when) " +
               "values(" + nzp_kvar + ", 5, to_date('01.04.2016','dd.mm.yyyy'), to_date('01.01.3000','dd.mm.yyyy'), '" + prm5 + "', 1, 1, to_date('01.04.2016','dd.mm.yyyy'))";
            cmd = new NpgsqlCommand(cmdText, conn);
            cmd.ExecuteNonQuery();
            conn.Close();
            return 1;
        }

        public DataTable SelectGkhCode(string ulica, string dom)
        {
            connStr = "Server=85.140.61.250;Database=gkh_samara;User ID=bars;Password=md5SM3tv;CommandTimeout=180000;";
            //string cmdText = "SELECT gro.gkh_code, gro.address from GKH_REALITY_OBJECT gro where replace(lower(GRO.ADDRESS), ' ','') LIKE replace(lower('%" + ulica + "%'), ' ','') " +
            //    //" and MUNICIPALITY_ID in (21690, 21691, 21692, 21693, 21694, 21695, 21696, 21697, 21698) and replace(lower(GRO.ADDRESS), ' ','') LIKE replace(lower('%д. " + dom + "'), ' ','')";
            //    " and MUNICIPALITY_ID in (21686) and replace(lower(GRO.ADDRESS), ' ','') LIKE replace(lower('%д. " + dom + "'), ' ','')";
            string cmdText = "SELECT gro.gkh_code, gro.address from GKH_REALITY_OBJECT gro where replace(lower(GRO.ADDRESS), ' ','') LIKE replace(lower('%" + ulica + "'), ' ','') " +
               " and MUNICIPALITY_ID in (21690, 21691, 21692, 21693, 21694, 21695, 21696, 21697, 21698)";
                        
            //string cmdText = "SELECT id FROM realty_object where mu_name LIKE '%Волжский р-н%'";
            //string cmdText = "SELECT * FROM employees";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            //cmd.Parameters.Add("surname", surname);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            try
            {
                da.Fill(dt);
                if (dt.Rows.Count == 1)
                    return dt;
                else
                    return null;

            }
            catch (Exception e)
            {
                return null;
            }
        }

        public int UpdatePkod(string nzp_kvar, string pkod)
        {
            string connStr = "Server=192.168.1.25;Database=billAukNew;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "SELECT nzp_kvar FROM bill01_data.kvar where nzp_kvar = " + nzp_kvar;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            if (dt.Rows.Count != 1)
                return 0;
            int nzp_kvar1 = Convert.ToInt32(dt.Rows[0][0]);
            cmdText = "UPDATE bill01_data.kvar SET pkod = " + pkod + " where nzp_kvar = " + nzp_kvar1;
            cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            cmdText = "UPDATE fbill_data.kvar SET pkod = " + pkod + " where nzp_kvar = " + nzp_kvar1;
            cmd = new NpgsqlCommand(cmdText, conn);
            cmd.ExecuteNonQuery();
            conn.Close();
            return 1;
        }

        public void InsertRoomCount(string nkvar, string num_ls, string fio, string val)
        {
            string connStr = "Server=192.168.1.25;Database=billDemidov;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "SELECT nzp_kvar FROM bill01_data.kvar where nzp_dom = 7155109 AND nkvar = '" + nkvar + "' and num_ls = " + num_ls;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            int nzp_kvar = Convert.ToInt32(dt.Rows[0][0]);
            cmdText = "INSERT INTO bill01_data.prm_1(nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user, dat_when) " +
                "values(" + nzp_kvar + ", 107, to_date('01.07.2015','dd.mm.yyyy'), to_date('01.01.3000','dd.mm.yyyy'), '" + val + "', 1, 1, to_date('27.07.2015','dd.mm.yyyy'))";
            Console.WriteLine(cmdText);
            cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        public void InsertRoomCount(string nzpKvar, string val)
        {
            string connStr = "Server=192.168.1.25;Database=billDemidov;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = "INSERT INTO bill01_data.prm_1(nzp, nzp_prm, dat_s, dat_po, val_prm, is_actual, nzp_user, dat_when, month_calc) " +
                "VALUES(" + nzpKvar + ", 107, to_date('2015-04-01','yyyy-mm-dd'), to_date('3000-01-31','yyyy-mm-dd'), '" + val + "', 1, 133, to_date('2015-05-27','yyyy-mm-dd'), to_date('2015-04-01','yyyy-mm-dd'))";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            cmd.ExecuteNonQuery();
            conn.Close();
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

        public DataTable GetMinusPeni()
        {
            string localhost = "";
            string database = "billAuk";
            //string connStr = "Server=localhost;Database=billTlt;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string connStr = "Server=" + (localhost == "localhost" ? "localhost" : "192.168.1.25") + ";Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = @"SELECT k.nzp_kvar, sum(sum_outsaldo)
                                FROM bill01_charge_16.charge_03 c
                                INNER JOIN bill01_data.kvar k on k.nzp_kvar = c.nzp_kvar 
                                INNER JOIN bill01_data.dom d on d.nzp_dom = k.nzp_dom
                                INNER JOIN bill01_data.s_ulica ul on ul.nzp_ul = d.nzp_ul
                                WHERE nzp_serv = 500 and dat_charge is null
                                group by 1
                                having sum(sum_outsaldo) <= -100";
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

        public DataTable GetSaldoAndParam()
        {
            string localhost = "";
            string database = "billTlt";
            //string connStr = "Server=localhost;Database=billTlt;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string connStr = "Server=" + (localhost == "localhost" ? "localhost" : "192.168.1.25") + ";Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = @"SELECT ul.ulica, d.ndom, k.nkvar, k.nkvar_n, k.num_ls, p107.val_prm as room_count, p4.val_prm as total_area, p6.val_prm as living_area, 
                            CASE WHEN p8.val_prm = '1' THEN 'Да' ELSE 'Нет' END as privatiz,
                            CASE WHEN p3.val_prm = '1' THEN 'изолированная' ELSE 'коммунальная' END as type_,
                            k.fio, p1010270.val_prm, p5.val_prm, coalesce(sum(c.sum_insaldo), 0) - coalesce(sum(c.sum_money), 0) as sum_insaldo, peni.sum_insaldo_peni
                            FROM  bill01_data.kvar k
                            LEFT JOIN bill01_charge_16.charge_03 c on k.nzp_kvar = c.nzp_kvar and c.nzp_serv not in (1,500) and dat_charge is null
                            INNER JOIN bill01_data.dom d on d.nzp_dom = k.nzp_dom
                            INNER JOIN bill01_data.s_ulica ul on ul.nzp_ul = d.nzp_ul
                            LEFT JOIN (SELECT nzp, max(val_prm) as val_prm from bill01_data.prm_1 where is_actual = 1 AND Extract(year from dat_po) = 3000 AND nzp_prm = 107 group by 1) p107 on p107.nzp = k.nzp_kvar
                            LEFT JOIN (SELECT nzp, max(val_prm) as val_prm from bill01_data.prm_1 where is_actual = 1 AND Extract(year from dat_po) = 3000 AND nzp_prm = 4 group by 1) p4 on p4.nzp = k.nzp_kvar
                            LEFT JOIN (SELECT nzp, max(val_prm) as val_prm from bill01_data.prm_1 where is_actual = 1 AND Extract(year from dat_po) = 3000 AND nzp_prm = 6 group by 1) p6 on p6.nzp = k.nzp_kvar
                            LEFT JOIN (SELECT nzp, max(val_prm) as val_prm from bill01_data.prm_1 where is_actual = 1 AND Extract(year from dat_po) = 3000 AND nzp_prm = 8 group by 1) p8 on p8.nzp = k.nzp_kvar
                            LEFT JOIN (SELECT nzp, max(val_prm) as val_prm from bill01_data.prm_1 where is_actual = 1 AND Extract(year from dat_po) = 3000 AND nzp_prm = 3 group by 1) p3 on p3.nzp = k.nzp_kvar
                            LEFT JOIN (SELECT nzp, max(val_prm) as val_prm from bill01_data.prm_1 where is_actual = 1 AND Extract(year from dat_po) = 3000 AND nzp_prm = 1010270 group by 1) p1010270 on p1010270.nzp = k.nzp_kvar
                            LEFT JOIN (SELECT nzp, max(val_prm) as val_prm from bill01_data.prm_1 where is_actual = 1 AND Extract(year from dat_po) = 3000 AND nzp_prm = 5 group by 1) p5 on p5.nzp = k.nzp_kvar
                            LEFT JOIN (SELECT coalesce(sum(sum_insaldo), 0) - coalesce(sum(sum_money), 0) as sum_insaldo_peni, nzp_kvar FROM bill01_charge_16.charge_03 where nzp_serv = 500 and dat_charge is null group by 2) peni on peni.nzp_kvar = k.nzp_kvar
                            group by 1,2,3,4,5,6,7,8,9,10,11,12,13,15
                            ORDER BY 1,2,3";
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

        public DataTable GetSaldoAndParamByServ()
        {
            string localhost = "";
            string database = "billTlt";
            //string connStr = "Server=localhost;Database=billTlt;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string connStr = "Server=" + (localhost == "localhost" ? "localhost" : "192.168.1.25") + ";Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = @"SELECT ul.ulica, d.ndom, k.nkvar, k.nkvar_n, k.num_ls, p107.val_prm as room_count, p4.val_prm as total_area, p6.val_prm as living_area, 
                            CASE WHEN p8.val_prm = '1' THEN 'Да' ELSE 'Нет' END as privatiz,
                            CASE WHEN p3.val_prm = '1' THEN 'изолированная' ELSE 'коммунальная' END as type_,
                            k.fio, p1010270.val_prm, p5.val_prm, coalesce(sum(c.sum_insaldo), 0) - coalesce(sum(c.sum_money), 0) as sum_insaldo, s.service
                            FROM  bill01_data.kvar k
                            LEFT JOIN bill01_charge_16.charge_03 c on k.nzp_kvar = c.nzp_kvar and c.nzp_serv not in (1) and dat_charge is null
                            LEFT JOIN fbill_kernel.services s on s.nzp_serv = c.nzp_serv
                            INNER JOIN bill01_data.dom d on d.nzp_dom = k.nzp_dom
                            INNER JOIN bill01_data.s_ulica ul on ul.nzp_ul = d.nzp_ul
                            LEFT JOIN (SELECT nzp, max(val_prm) as val_prm from bill01_data.prm_1 where is_actual = 1 AND Extract(year from dat_po) = 3000 AND nzp_prm = 107 group by 1) p107 on p107.nzp = k.nzp_kvar
                            LEFT JOIN (SELECT nzp, max(val_prm) as val_prm from bill01_data.prm_1 where is_actual = 1 AND Extract(year from dat_po) = 3000 AND nzp_prm = 4 group by 1) p4 on p4.nzp = k.nzp_kvar
                            LEFT JOIN (SELECT nzp, max(val_prm) as val_prm from bill01_data.prm_1 where is_actual = 1 AND Extract(year from dat_po) = 3000 AND nzp_prm = 6 group by 1) p6 on p6.nzp = k.nzp_kvar
                            LEFT JOIN (SELECT nzp, max(val_prm) as val_prm from bill01_data.prm_1 where is_actual = 1 AND Extract(year from dat_po) = 3000 AND nzp_prm = 8 group by 1) p8 on p8.nzp = k.nzp_kvar
                            LEFT JOIN (SELECT nzp, max(val_prm) as val_prm from bill01_data.prm_1 where is_actual = 1 AND Extract(year from dat_po) = 3000 AND nzp_prm = 3 group by 1) p3 on p3.nzp = k.nzp_kvar
                            LEFT JOIN (SELECT nzp, max(val_prm) as val_prm from bill01_data.prm_1 where is_actual = 1 AND Extract(year from dat_po) = 3000 AND nzp_prm = 1010270 group by 1) p1010270 on p1010270.nzp = k.nzp_kvar
                            LEFT JOIN (SELECT nzp, max(val_prm) as val_prm from bill01_data.prm_1 where is_actual = 1 AND Extract(year from dat_po) = 3000 AND nzp_prm = 5 group by 1) p5 on p5.nzp = k.nzp_kvar
                            group by 1,2,3,4,5,6,7,8,9,10,11,12,13,15
                            ORDER BY 1,2,3";
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

        public DataTable GetTarifDomofon()
        {
            string localhost = "";
            string database = "billTlt";
            //string connStr = "Server=localhost;Database=billTlt;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string connStr = "Server=" + (localhost == "localhost" ? "localhost" : "192.168.1.25") + ";Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = @"SELECT ul.ulica, d.ndom, k.nkvar, k.nkvar_n, k.num_ls, c.rsum_tarif
                                FROM  bill01_data.kvar k
                                LEFT JOIN bill01_charge_16.charge_03 c on k.nzp_kvar = c.nzp_kvar and c.nzp_serv in (26) and dat_charge is null
                                INNER JOIN bill01_data.dom d on d.nzp_dom = k.nzp_dom
                                INNER JOIN bill01_data.s_ulica ul on ul.nzp_ul = d.nzp_ul
                                ORDER BY 1,2,3";
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

        public DataTable GetServList()
        {
            string localhost = "";
            string database = "billTlt";
            //string connStr = "Server=localhost;Database=billTlt;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string connStr = "Server=" + (localhost == "localhost" ? "localhost" : "192.168.1.25") + ";Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = @"SELECT ul.ulica, d.ndom, k.nkvar, k.nkvar_n, k.num_ls, serv.service,s.name_supp, t.dat_s, t.dat_po
                                FROM  bill01_data.kvar k
                                INNER JOIN bill01_data.dom d on d.nzp_dom = k.nzp_dom
                                INNER JOIN bill01_data.s_ulica ul on ul.nzp_ul = d.nzp_ul
                                LEFT JOIN bill01_data.tarif t on t.nzp_kvar = k.nzp_kvar
                                left outer join bill01_kernel.supplier s on t.nzp_supp = s.nzp_supp 
				                left join fbill_kernel.services serv on serv.nzp_serv = t.nzp_serv                               
                                ORDER BY 1,2,3";
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

        public DataTable GetCountersVal(string data)
        {
            string connStr = "Server=192.168.1.25;Database=billTlt;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);

            int month = Convert.ToInt32(data.Split('.')[0]);
            int year = Convert.ToInt32(data.Split('.')[1]);
            int days = DateTime.DaysInMonth(year, month);

            int monthPred = month - 1;
            int yearPred = year;
            if (monthPred == 0)
            {
                monthPred = 12;
                yearPred--;
            }
            int daysPred = DateTime.DaysInMonth(yearPred, monthPred);

            conn.Open();
            try
            {
                string cmdText = @"drop table if exists t_couns_test";
                NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
                cmd.ExecuteNonQuery();

                cmdText = @"Create temp table t_couns_test(nzp_kvar integer, nzp_serv INTEGER, service char(30), 
                            num_cnt char(30), cnt_stage INTEGER, mmnog numeric(16,7), dat_install Date, dat_uchet_pred Date, val_cnt_pred numeric(14,2), dat_uchet Date, val_cnt numeric(14,2))";
                cmd = new NpgsqlCommand(cmdText, conn);
                cmd.ExecuteNonQuery();

                cmdText = @"INSERT INTO t_couns_test
                            SELECT cs.nzp_kvar as nzp_kvar, cs.nzp_serv as nzp_serv, serv.service as service, cs1.num_cnt as num_cnt, sc.cnt_stage as cnt_stage, sc.mmnog as mmnog, 
                            dat_install as dat_install, null as dat_uchet_pred, 0 as val_cnt_pred, max(dat_uchet) as dat_uchet, 0 as val_cnt
                            FROM bill01_data.counters cs
                            INNER JOIN bill01_data.counters_spis cs1 on cs1.nzp_counter = cs.nzp_counter
                            INNER JOIN fbill_data.kvar k on k.nzp_kvar = cs.nzp_kvar
                            INNER JOIN fbill_data.dom d on d.nzp_dom = k.nzp_dom 
                            INNER JOIN fbill_data.s_ulica u on u.nzp_ul = d.nzp_ul 
                            INNER JOIN fbill_kernel.services serv on serv.nzp_serv = cs.nzp_serv
                            LEFT JOIN bill01_kernel.s_counttypes sc on cs1.nzp_cnttype = sc.nzp_cnttype
                            LEFT JOIN (SELECT nzp, max(val_prm::date) as dat_install FROM bill01_data.prm_17 where nzp_prm = 2025 and is_actual = 1 and extract(year from dat_po) = 3000 group by 1) p17 on p17.nzp = cs.nzp_counter
                            where cs.is_actual != 100 and dat_uchet >= '" + year + "-" + month.ToString("00") + "-01'::date and dat_uchet <= '" + 
                            year + "-" + month.ToString("00") + "-" + days.ToString("00") + "'::date group by 1,2,3,4,5,6,7";
                cmd = new NpgsqlCommand(cmdText, conn);
                cmd.ExecuteNonQuery();

                cmdText = @"UPDATE t_couns_test SET val_cnt = 
                            (SELECT max(val_cnt) 
                            FROM bill01_data.counters c 
                            inner join bill01_data.counters_spis cs on cs.nzp_counter = c.nzp_counter 
                            where c.nzp_kvar = t_couns_test.nzp_kvar and c.nzp_serv = t_couns_test.nzp_serv and cs.num_cnt = t_couns_test.num_cnt and c.dat_uchet = t_couns_test.dat_uchet 
                                and cs.is_actual != 100)";
                cmd = new NpgsqlCommand(cmdText, conn);
                cmd.ExecuteNonQuery();

                cmdText = @"UPDATE t_couns_test SET dat_uchet_pred = 
                            (SELECT max(dat_uchet) 
                            FROM bill01_data.counters c 
                            inner join bill01_data.counters_spis cs on cs.nzp_counter = c.nzp_counter 
                            where c.nzp_kvar = t_couns_test.nzp_kvar and c.nzp_serv = t_couns_test.nzp_serv and cs.num_cnt = t_couns_test.num_cnt 
                            and dat_uchet >= '" + yearPred + "-" + monthPred.ToString("00") + @"-01'::date and dat_uchet <= '" + 
                            yearPred + "-" + monthPred.ToString("00") + "-" + daysPred.ToString("00") + "'::date and cs.is_actual != 100)";
                cmd = new NpgsqlCommand(cmdText, conn);
                cmd.ExecuteNonQuery();

                cmdText = @"UPDATE t_couns_test SET val_cnt_pred = 
                            (SELECT max(val_cnt) 
                            FROM bill01_data.counters c 
                            inner join bill01_data.counters_spis cs on cs.nzp_counter = c.nzp_counter 
                            where c.nzp_kvar = t_couns_test.nzp_kvar and c.nzp_serv = t_couns_test.nzp_serv and cs.num_cnt = t_couns_test.num_cnt 
                                and c.dat_uchet = t_couns_test.dat_uchet_pred and cs.is_actual != 100)";
                cmd = new NpgsqlCommand(cmdText, conn);
                cmd.ExecuteNonQuery();

                cmdText = @"SELECT ul.ulica, d.ndom, k.nkvar, t.service, t.num_cnt, t.cnt_stage, t.mmnog, t.dat_install, t.dat_uchet_pred, t.val_cnt_pred, t.dat_uchet, t.val_cnt
                            FROM t_couns_test t
                            INNER JOIN bill01_data.kvar k on k.nzp_kvar = t.nzp_kvar
                            INNER JOIN bill01_data.dom d on d.nzp_dom = k.nzp_dom
                            INNER JOIN bill01_data.s_ulica ul on ul.nzp_ul = d.nzp_ul";
                cmd = new NpgsqlCommand(cmdText, conn);
                NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                return dt;
            }
            catch
            {
                return null;
            }
            finally
            {
                conn.Close();
            }
        }

        public DataTable GetPeniSuppByNzpKvar(string nzpKvar)
        {
            string localhost = "";
            string database = "billAuk";
            //string connStr = "Server=localhost;Database=billTlt;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string connStr = "Server=" + (localhost == "localhost" ? "localhost" : "192.168.1.25") + ";Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = @"SELECT nzp_kvar, sum_outsaldo, num_ls, nzp_supp FROM bill01_charge_16.charge_03 where nzp_kvar = " + nzpKvar + " and nzp_serv = 500 and dat_charge is null";
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

        public DataTable GetRsumTarifSuppAndServByNzpKvar(string nzpKvar)
        {
            string localhost = "";
            string database = "billAuk";
            //string connStr = "Server=localhost;Database=billTlt;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string connStr = "Server=" + (localhost == "localhost" ? "localhost" : "192.168.1.25") + ";Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = @"SELECT nzp_kvar, rsum_tarif, num_ls, nzp_serv, nzp_supp FROM bill01_charge_16.charge_03 where nzp_kvar = " + nzpKvar + " and nzp_serv not in (1, 500) and dat_charge is null and rsum_tarif > 0";
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

        public DataTable GetFirstRsumTarifSuppAndServByNzpKvar(string nzpKvar)
        {
            string localhost = "";
            string database = "billAuk";
            //string connStr = "Server=localhost;Database=billTlt;User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string connStr = "Server=" + (localhost == "localhost" ? "localhost" : "192.168.1.25") + ";Database=" + database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            string cmdText = @"SELECT nzp_kvar, rsum_tarif, num_ls, nzp_serv, nzp_supp FROM bill01_charge_16.charge_03 where nzp_kvar = " + nzpKvar + 
                            " and nzp_serv not in (1, 500) and dat_charge is null LIMIT 1";
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
    }
}
