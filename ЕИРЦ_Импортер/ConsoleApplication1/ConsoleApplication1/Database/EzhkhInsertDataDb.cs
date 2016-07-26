using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Npgsql;
using System.Data;

namespace ConsoleApplication1.Database
{
    class EzhkhInsertDataDb
    {
        private string connStr;

        public EzhkhInsertDataDb()
        {
            connStr = "Server=85.140.61.250;Database=gkh_samara;User ID=bars;Password=md5SM3tv;CommandTimeout=180000;";
        }

        public string InsertHouseManOrg(string gkh_code, string date, string manOrgId)
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
            string cmdText1 = "INSERT INTO gkh_morg_contract (ID, object_version, object_create_date, object_edit_date, manag_org_id, type_contract, start_date" +
                    ") VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, '" + manOrgId + "', 20" +
                    ",TO_TIMESTAMP('" + date + "', 'DD.MM.YYYY')) RETURNING id";
            NpgsqlCommand cmd1 = new NpgsqlCommand(cmdText1, conn);
            
            try
            {
                int id2 = Convert.ToInt32(cmd1.ExecuteScalar());
                string cmdText2 = "INSERT INTO gkh_morg_contract_realobj (id, object_version, object_create_date, object_edit_date, reality_obj_id, man_org_contract_id" +
                    ") VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, '" + id +
                    "','" + id2 + "')";
                NpgsqlCommand cmd2 = new NpgsqlCommand(cmdText2, conn);
                cmd2.ExecuteNonQuery();

                string cmdText3 = "INSERT INTO gkh_morg_contract_owners (id, object_version, object_create_date, object_edit_date, contract_foundation" +
                    ") VALUES(" + id2 + ", 0, CURRENT_DATE, CURRENT_DATE, 20)";
                NpgsqlCommand cmd3 = new NpgsqlCommand(cmdText3, conn);
                cmd3.ExecuteNonQuery();
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

        public string InsertCommunalOrg(string manOrgId)
        {
            string cmdText = "select realityobject_id from gkh_supply_resorg_ro where supply_resorg_id = " + manOrgId;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            try
            {
                da.Fill(dt);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string cmdText1 = "INSERT INTO gkh_obj_resorg (ID, object_version, object_create_date, object_edit_date, resorg_id, reality_object_id, date_start" +
                       ") VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, '" + manOrgId + "'," + dt.Rows[i][0].ToString() + "," +
                       "TO_TIMESTAMP('01.01.2014', 'DD.MM.YYYY'))";
                    NpgsqlCommand cmd2 = new NpgsqlCommand(cmdText1, conn);
                    cmd2.ExecuteNonQuery();
                }
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

        public string InsertResOrg(string manOrgId)
        {
            string cmdText = "select realityobject_id from gkh_public_servorg_ro where public_servorg_id = " + manOrgId;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            try
            {
                da.Fill(dt);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string cmdText1 = "INSERT INTO gkh_ro_pub_servorg (ID, object_version, object_create_date, object_edit_date, pub_servorg_id, real_obj_id, date_start" +
                       ") VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, '" + manOrgId + "'," + dt.Rows[i][0].ToString() + "," +
                       "TO_TIMESTAMP('01.09.2014', 'DD.MM.YYYY'))";
                    NpgsqlCommand cmd2 = new NpgsqlCommand(cmdText1, conn);
                    cmd2.ExecuteNonQuery();
                }
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

        public string UpdateMkdArea(int id, decimal value)
        {
            connStr = "Server=85.140.61.250;Database=gkh_samara;User ID=bars;Password=md5SM3tv;CommandTimeout=180000;";
            string cmdText = "UPDATE gkh_reality_object SET area_mkd = '" + value + "' where id = " + id;
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            try
            {
                cmd.ExecuteNonQuery();
                return "ЗАГРУЖЕНО";
            }
            catch(Exception e)
            {
                string err = e.Message;
                return err;
            }
            finally
            {
                conn.Close();
            }
            
        }

        public string UpdatePhysicalWear(string gkh_code, decimal value)
        {
            connStr = "Server=85.140.61.250;Database=gkh_samara;User ID=bars;Password=md5SM3tv;CommandTimeout=180000;";
            string cmdText = "UPDATE gkh_reality_object SET physical_wear = " + value + " where gkh_code = '" + gkh_code + "'";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            try
            {
                cmd.ExecuteNonQuery();
                return "ЗАГРУЖЕНО";
            }
            catch (Exception e)
            {
                string err = e.Message;
                return err;
            }
            finally
            {
                conn.Close();
            }

        }

        public string UpdatePhysicalWearTehPassport(string gkh_code, decimal value)
        {
            connStr = "Server=85.140.61.250;Database=gkh_samara;User ID=bars;Password=md5SM3tv;CommandTimeout=180000;";
            string cmdText = @"UPDATE tp_teh_passport_value set value = '"+ value + 
                @"'WHERE form_code = 'Form_1' and cell_code = '20:1' and teh_passport_id in 
                    (SELECT id FROM tp_teh_passport where reality_obj_id in (SELECT id FROM gkh_reality_object where  gkh_code = '"+ gkh_code + "'))";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            try
            {
                int i = cmd.ExecuteNonQuery();
                if(i == 0)
                {
                    cmdText = @"INSERT INTO tp_teh_passport_value(object_version, object_create_date, object_edit_date, teh_passport_id, form_code, cell_code, value)
                                SELECT 0, current_date, current_date, id, 'Form_1', '20:1', '"+ value + 
                                "' FROM tp_teh_passport where reality_obj_id in (SELECT id FROM gkh_reality_object where  gkh_code = '"+ gkh_code + "')";
                    cmd = new NpgsqlCommand(cmdText, conn);
                    i = cmd.ExecuteNonQuery();
                    return "ЗАГРУЖЕНО|" + i;
                }
                    
                return "ЗАГРУЖЕНО|" + i;
            }
            catch (Exception e)
            {
                string err = e.Message;
                return "НЕЗАГРУЖЕНО|" + err;
            }
            finally
            {
                conn.Close();
            }

        }
    }
}
