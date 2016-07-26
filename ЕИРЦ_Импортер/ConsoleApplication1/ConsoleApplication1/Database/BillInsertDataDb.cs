using Npgsql;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1.Database
{
    class BillInsertDataDb
    {
        private string connStr;
        private string connStrTest;
        private string connStrHome;
        public BillInsertDataDb(string database)
        {
            //connStr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.126.128)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=XE)));User Id = HR; Password = test";
            //connStr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.1.30)(PORT=1521))(CONNECT_DATA=(SID=ORCL)));User Id=dbq;Password=dbq;";
            //connStr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.5.77)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=orcl)));User Id = MJF; Password = ActaNonVerba";
            connStr = "Server=192.168.1.25;Database="+ database + ";User ID=postgres;Password=Admin;CommandTimeout=180000;";
            connStrTest = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=85.140.61.250)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ezhkh)));User Id = b4_gkh_samara2; Password = ACTANONVERBA";
            connStrHome = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=46.0.13.2)(PORT=1578))(CONNECT_DATA=(SID=orcl)));User Id=dbq;Password=dbq";
        }

        public void UpdateKvarTotalSquare(string val_prm, string nzp)
        {
            Decimal value = Math.Round(Convert.ToDecimal(val_prm), 2);
            string cmdText = @"UPDATE bill01_data.prm_1 set val_prm = '" + value +
                @"' WHERE nzp = " + nzp + " AND is_actual = 1 AND extract(year from dat_po) = 3000 AND nzp_prm = 4";
            NpgsqlConnection conn = new NpgsqlConnection(connStr);
            NpgsqlCommand cmd = new NpgsqlCommand(cmdText, conn);
            conn.Open();
            try
            {
                int i = cmd.ExecuteNonQuery();
                if (i == 0)
                {
                    cmdText = @"INSERT INTO bill01_data.prm_1(nzp,nzp_prm,dat_s,dat_po,val_prm,is_actual,cur_unl,nzp_user,dat_when) 
                                VALUES("+ nzp + ", 4, '2016-07-01', '3000-01-01', '"+ value + "', 1, 1, 1, current_date)";
                    cmd = new NpgsqlCommand(cmdText, conn);
                    i = cmd.ExecuteNonQuery();
                }
            }
            catch (Exception e)
            {
                string err = e.Message;
            }
            finally
            {
                conn.Close();
            }
        }
    }
}
