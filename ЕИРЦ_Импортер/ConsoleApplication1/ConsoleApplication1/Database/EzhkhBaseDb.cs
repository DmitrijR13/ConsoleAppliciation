using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Npgsql;
using System.Data;

namespace ConsoleApplication1.Database
{
    class EzhkhBaseDb
    {
        private string connStr;

        public EzhkhBaseDb()
        {
            connStr = "Server=85.140.61.250;Database=gkh_samara;User ID=bars;Password=md5SM3tv;CommandTimeout=180000;";
        }

        public Int32 SelectRoIDByGkhCode(string gkh_code)
        {
            string cmdText = "SELECT id from GKH_REALITY_OBJECT where gkh_code = " + gkh_code;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            int id;
            da.Fill(dt);
            if(dt.Rows.Count == 1)
                return Convert.ToInt32(dt.Rows[0][0]);
            else if(dt.Rows.Count == 0)
                return 0;
            else
                return 2;
        }

        public void DelCurRepair(Int32 roId)
        {
            string cmdText = @"delete FROM gkh_obj_curent_repair where reality_object_id = " + roId + " and (EXTRACT(year FROM fact_date) = 2016 OR EXTRACT(year FROM plan_date) = 2016 )";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd1 = new NpgsqlCommand(cmdText, conn);

            conn.Open();
            try
            {
                cmd1.ExecuteNonQuery();
            }
            catch (Exception e)
            {

            }
            finally
            { conn.Close(); }

        }

        public void InsertCurRepair(Int32 roId)
        {
            string cmdText = @"delete FROM gkh_obj_curent_repair where reality_object_id = " + roId + " and (EXTRACT(year FROM fact_date) = 2016 OR EXTRACT(year FROM plan_date) = 2016 )";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd1 = new NpgsqlCommand(cmdText, conn);

            conn.Open();
            try
            {
                cmd1.ExecuteNonQuery();
            }
            catch (Exception e)
            {

            }
            finally
            { conn.Close(); }

        }

        public List<String> SelectCurRepWorkId(String name)
        {
            List<String> data = new List<String>();
            string cmdText = @"SELECT cur.id, unit.name FROM gkh_dict_work_cur_repair cur INNER JOIN gkh_dict_unitmeasure unit on unit.id = cur.unit_measure_id 
                                where cur.name = '"+ name + "'";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            int id;
            da.Fill(dt);
            if (dt.Rows.Count == 1)
            {
                data.Add(Convert.ToString(dt.Rows[0][0]));
                data.Add(Convert.ToString(dt.Rows[0][1]));
                return data;
            }
            else
            {
                return new List<String>();
            }
        }

        public void InsertCurRepair(Int32 roId, String factDate, String factSum, String factWork, String planDate, String planSum, String planWork, String UnitMeasure, String workKindId)
        {
            string cmdText = "";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            try
            {
                    string cmdText1 = "INSERT INTO gkh_obj_curent_repair(id, object_version, object_create_date, object_edit_date, reality_object_id, fact_date, fact_sum, " +
                    " fact_work, plan_date, plan_sum, plan_work, unit_measure, work_kind_id) " + 
                    " VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, " + roId + ", TO_TIMESTAMP('" + factDate + "', 'YYYY-MM-DD')," + factSum + ", " + factWork +
                    ", TO_TIMESTAMP('" + planDate + "', 'YYYY-MM-DD'), " + planSum + ", "+ planWork + ", '"+ UnitMeasure + "', "+ workKindId + ")";
                    NpgsqlCommand cmd2 = new NpgsqlCommand(cmdText1, conn);
                    cmd2.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                
            }
            finally
            { conn.Close(); }
        }
    }
}
