using Npgsql;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1.Database
{
    class InsertPeopleDb
    {
        private string connStr;
        public InsertPeopleDb()
        {
            //connStr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.126.128)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=XE)));User Id = HR; Password = test";
            connStr = "Server=85.140.61.250;Database=gkh_samara;User ID=bars;Password=md5SM3tv;CommandTimeout=180000;";
            //connStr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.5.77)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=orcl)));User Id = MJF; Password = ActaNonVerba";
        }

        public List<string> houses = new List<string>();

        public string InsertPeople5(string gkh_code, Int32 flat, string total_area, string useful_area, string privatized, string residents_count, string fio)
        {
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

        public string InsertPeople(string gkh_code, string flat, string fio, string area)
        {
            string cmdText = "SELECT id from GKH_REALITY_OBJECT where gkh_code = " + gkh_code;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            int id;
            conn.Open();
            da.Fill(dt);
            id = Convert.ToInt32(dt.Rows[0][0]);
            if (flat.Substring(0, 1) == "0")
                flat = flat.Substring(1);
            if (flat.Substring(0, 1) == "0")
                flat = flat.Substring(1);
            string useful_area = "0";
            string total_area = "0";
            int residents_count = 0;
            int priv = 10;
            if (area != "не начисл." && area != "")
                total_area = area.Replace(',', '.');
            string cmdText1 = "INSERT INTO gkh_obj_apartment_info (id, object_version, object_create_date, object_edit_date, num_apartment, area_total, count_people, privatized," +
                    "reality_object_id, fio_owner, area_living) VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, '" + flat + "'," + total_area
                    + "," + residents_count + "," + priv + "," + id + ",'" + fio + "'," + useful_area + ")";
            NpgsqlCommand cmd1 = new NpgsqlCommand(cmdText1, conn);
            try
            {
                cmd1.ExecuteNonQuery();
                return "ЗАГРУЖЕНО";
            }
            catch (Exception e)
            {
                string err = e.Message;
                return err;
            }
            finally
            { conn.Close(); }

        }

        public string InsertPeople(string gkh_code, string flat, string fio,
            string useful_area, string total_area, string residents_count)
        {
            string cmdText = "SELECT id from GKH_REALITY_OBJECT where gkh_code = " + gkh_code;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            int id;
            conn.Open();
            try
            {
                da.Fill(dt);
                id = Convert.ToInt32(dt.Rows[0][0]);
                if (flat.Substring(0, 1) == "0")
                    flat = flat.Substring(1);
                if (flat.Substring(0, 1) == "0")
                    flat = flat.Substring(1);
                string cmdCheck = "SELECT id from gkh_obj_apartment_info where reality_object_id = " + id +
                    " and num_apartment = '" + flat + "'";
                NpgsqlCommand cmd2 = new NpgsqlCommand(cmdCheck, conn);
                NpgsqlDataAdapter da2 = new NpgsqlDataAdapter(cmd2);
                DataTable dt2 = new DataTable();
                int id2;
                da2.Fill(dt2);
                id2 = Convert.ToInt32(dt2.Rows[0][0]);
                if (useful_area == null || useful_area == "" || useful_area == " ")
                    useful_area = "0";
                string cmdText1 = "UPDATE gkh_obj_apartment_info set area_total = " + total_area + ", count_people = " + residents_count + ", " +
                        "fio_owner = '" + fio + "', area_living = " + useful_area + " where id = " + id2;
                NpgsqlCommand cmd1 = new NpgsqlCommand(cmdText1, conn);

                cmd1.ExecuteNonQuery();
                return "ЗАГРУЖЕНО";
            }
            catch (Exception e)
            {
                string err = e.Message;
                return err + "|" + flat;
            }
            finally
            { conn.Close(); }

        }

        public string InsertPeople(string gkh_code, string flat, string fio,
            string useful_area, string total_area, string residents_count, string privatized)
        {
            string cmdText = "SELECT id from GKH_REALITY_OBJECT where gkh_code = " + gkh_code;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            int id;
            conn.Open();
            da.Fill(dt);
            id = Convert.ToInt32(dt.Rows[0][0]);
            if (flat.Substring(0, 1) == "0")
                flat = flat.Substring(1);
            if (flat.Substring(0, 1) == "0")
                flat = flat.Substring(1);
            int priv = 30;
            if (privatized == "да.")
                priv = 10;
            else
                priv = 20;
            if (useful_area == null || useful_area == "" || useful_area == " ")
                useful_area = "0";
            if (total_area == null || total_area == "" || total_area == " ")
                total_area = "0";
            if (residents_count == null || residents_count == "" || residents_count == " ")
                residents_count = "0";
            string cmdText1 = "INSERT INTO gkh_obj_apartment_info (id, object_version, object_create_date, object_edit_date, num_apartment, area_total, count_people, privatized," +
                    "reality_object_id, fio_owner, area_living) VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, '" + flat + "'," + total_area
                    + "," + residents_count + "," + priv + "," + id + ",'" + fio + "'," + useful_area + ")";
            NpgsqlCommand cmd1 = new NpgsqlCommand(cmdText1, conn);
            try
            {
                cmd1.ExecuteNonQuery();
                return "ЗАГРУЖЕНО";
            }
            catch (Exception e)
            {
                string err = e.Message;
                return err;
            }
            finally
            { conn.Close(); }

        }
    }
}
