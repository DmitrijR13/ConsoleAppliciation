using Npgsql;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1.Database
{
    class GetRoIdentifier
    {
        public int SelectRoId(string municipalityId, string address)
        {
            string connStr = "Server=85.140.61.250;Database=gkh_samara;User ID=bars;Password=md5SM3tv;CommandTimeout=180000;";
            string cmdText = "SELECT gro.id from GKH_REALITY_OBJECT gro where replace(lower(GRO.ADDRESS), ' ','') LIKE replace(lower('%" + address + "'), ' ','') " +
               " and MUNICIPALITY_ID in ("+ municipalityId + ")";

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
                    return Convert.ToInt32(dt.Rows[0][0]);
                else
                    return 0;

            }
            catch (Exception e)
            {
                return 0;
            }
        }
    }
}
