using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
    class Dbf
    {
        #region конструкторы

        /// <summary>
        /// Конструктор
        /// </summary>
        public Dbf()
        {
            this.conn = new System.Data.Odbc.OdbcConnection();
            conn.ConnectionString = @"Driver={Microsoft dBase Driver (*.dbf)};datasource=dBase Files;" +
                                       "SourceType=DBF;Exclusive=No;" +
                                       "Collate=Machine;NULL=NO;DELETED=NO;" +
                                       "BACKGROUNDFETCH=NO;";

        }

        #endregion

        #region private fields

        private OdbcConnection conn = null;

        #endregion

        public DataTable SelectHouse(string db)
        {
            string strConn = "Provider=VFPOLEDB.1;Data Source=" + db + ";Collating Sequence=MACHINE";
            string strSelect = "SELECT  FROM kart";
            OleDbConnection myConn = new OleDbConnection(strConn);
            OleDbCommand myCommand = new OleDbCommand(strSelect, myConn);
            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myCommand);
            myConn.Open();
            DataTable dt = new DataTable();
            myDataAdapter.Fill(dt);
            return dt;
        }

        public DataTable SelectStreet(string db)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + db + ";Extended Properties=dBASE IV";
            string strSelect = "SELECT * FROM s_street";
            OleDbConnection myConn = new OleDbConnection(strConn);
            OleDbCommand myCommand = new OleDbCommand(strSelect, myConn);
            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myCommand);
            myConn.Open();
            DataTable dt = new DataTable();
            myDataAdapter.Fill(dt);
            return dt;

        }

        public void CreateDbf(string name)
        {
            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\temp;Extended Properties=dBase IV";

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            using (OleDbCommand command = connection.CreateCommand())
            {
                connection.Open();

                command.CommandText = "CREATE TABLE " + name.Replace('.', '_') + @" (Company Character(10), Date_PAY Date, PACCOUNT Character(15), COLD_1 Double, COLD_2 Double, ELEC_D Double,
                                        ELEC_N Double, HOT_1 Double, HOT_2 Double)";
                command.ExecuteNonQuery();
            }

        }

        public void InsertRow(string name, string pkod, DateTime date_pay, Decimal elec_d)
        {
            string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\temp;Extended Properties=dBase IV";

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            using (OleDbCommand command = connection.CreateCommand())
            {
                connection.Open();

                command.CommandText = "INSERT INTO " + name.Replace('.', '_') + @"(Company, Date_PAY, PACCOUNT, COLD_1, COLD_2, ELEC_D, ELEC_N, HOT_1, HOT_2) 
                                                                VALUES('', '" + date_pay.ToShortDateString() + "', " + pkod + ",0,0,"+elec_d+",0,0,0)";
                command.ExecuteNonQuery();
            }

        }

        public DataTable SelectLS(string db)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + db + ";Extended Properties=dBASE IV;Persist Security Info=False;";
            string strSelect = "SELECT kodin, o_area, p_area, kolgp, kolg FROM kart";
            OleDbConnection myConn = new OleDbConnection(strConn);
            OleDbCommand myCommand = new OleDbCommand(strSelect, myConn);
            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myCommand);
            myConn.Open();
            DataTable dt = new DataTable();
            myDataAdapter.Fill(dt);
            return dt;
        }

        public DataTable SelectKGLC(string db)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + db + ";Extended Properties=dBASE IV;Persist Security Info=False;";
            string strSelect = "SELECT * FROM kglc";
            OleDbConnection myConn = new OleDbConnection(strConn);
            OleDbCommand myCommand = new OleDbCommand(strSelect, myConn);
            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myCommand);
            myConn.Open();
            DataTable dt = new DataTable();
            myDataAdapter.Fill(dt);
            return dt;
        }

        public DataTable SelectPasskart(string db)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + db + ";Extended Properties=dBASE IV;Persist Security Info=False;";
            string strSelect = "SELECT geu, kkod FROM paskart";
            OleDbConnection myConn = new OleDbConnection(strConn);
            OleDbCommand myCommand = new OleDbCommand(strSelect, myConn);
            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myCommand);
            myConn.Open();
            DataTable dt = new DataTable();
            myDataAdapter.Fill(dt);
            return dt;
        }

        public void DeleteKglc(string db, string kkod)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + db + ";Extended Properties=dBASE IV;Persist Security Info=False;Mode=ReadWrite;";
            string strSelect = "DELETE FROM kglc k where k.kkod = " + kkod;
            OleDbConnection myConn = new OleDbConnection(strConn);
            OleDbCommand myCommand = new OleDbCommand(strSelect, myConn);
           
            myConn.Open();
            myCommand.ExecuteNonQuery();
            myConn.Close();
        }

        public DataTable SelectKartVU(string db)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + db + ";Extended Properties=dBASE IV;Persist Security Info=False;";
            string strSelect = "SELECT kkod, slv1 FROM kart_vu";
            OleDbConnection myConn = new OleDbConnection(strConn);
            OleDbCommand myCommand = new OleDbCommand(strSelect, myConn);
            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myCommand);
            myConn.Open();
            DataTable dt = new DataTable();
            myDataAdapter.Fill(dt);
            return dt;
        }

        public DataTable SelectKart(string db)
        {
            string strConn = "Provider=VFPOLEDB.1;Data Source=" + db + ";Collating Sequence=MACHINE";
            string strSelect = "SELECT * FROM kart";
            OleDbConnection myConn = new OleDbConnection(strConn);
            OleDbCommand myCommand = new OleDbCommand(strSelect, myConn);
            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myCommand);
            myConn.Open();
            DataTable dt = new DataTable();
            myDataAdapter.Fill(dt);
            return dt;
        }

        public DataTable SelectKodIN(string db)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + db + ";Extended Properties=dBASE IV;Persist Security Info=False;";
            string strSelect = "SELECT khouse, ktipb, skc FROM house_lit";
            OleDbConnection myConn = new OleDbConnection(strConn);
            OleDbCommand myCommand = new OleDbCommand(strSelect, myConn);
            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myCommand);
            myConn.Open();
            DataTable dt = new DataTable();
            myDataAdapter.Fill(dt);
            return dt;
        }

        public DataTable SelectHouse(string db, string kodIn)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + db + ";Extended Properties=dBASE IV;Persist Security Info=False;";
            string strSelect = "SELECT kod2,godp, kkod1, etag, geu, ud, kodin FROM s_house";
            OleDbConnection myConn = new OleDbConnection(strConn);
            OleDbCommand myCommand = new OleDbCommand(strSelect, myConn);
            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myCommand);
            myConn.Open();
            DataTable dt = new DataTable();
            myDataAdapter.Fill(dt);
            return dt;
        }

        public DataTable SelectNorm(string db)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + db + ";Extended Properties=dBASE IV";
            string strSelect = "SELECT VU, name, kod, sn FROM s_normtr";
            OleDbConnection myConn = new OleDbConnection(strConn);
            OleDbCommand myCommand = new OleDbCommand(strSelect, myConn);
            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myCommand);
            myConn.Open();
            DataTable dt = new DataTable();
            myDataAdapter.Fill(dt);
            return dt;

        }

        public DataTable SelectHouse2(string db)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + db + ";Extended Properties=dBASE IV;Persist Security Info=False;";
            string strSelect = "SELECT khouse, vu, mskt1, mskt2,mskt3 FROM house_vu";
            OleDbConnection myConn = new OleDbConnection(strConn);
            OleDbCommand myCommand = new OleDbCommand(strSelect, myConn);
            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myCommand);
            myConn.Open();
            DataTable dt = new DataTable();
            myDataAdapter.Fill(dt);
            return dt;
        }

        public DataTable SelectTarif(string db)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + db + ";Extended Properties=dBASE IV;Persist Security Info=False;";
            string strSelect = "SELECT krs, ktr, tipn FROM s_tarif";
            OleDbConnection myConn = new OleDbConnection(strConn);
            OleDbCommand myCommand = new OleDbCommand(strSelect, myConn);
            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myCommand);
            myConn.Open();
            DataTable dt = new DataTable();
            myDataAdapter.Fill(dt);
            return dt;
        }


        public DataTable SelectStreet(string db, string kodIn)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + db + ";Extended Properties=dBASE IV";
            string strSelect = "SELECT name, kodin FROM s_street";
            OleDbConnection myConn = new OleDbConnection(strConn);
            OleDbCommand myCommand = new OleDbCommand(strSelect, myConn);
            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myCommand);
            myConn.Open();
            DataTable dt = new DataTable();
            myDataAdapter.Fill(dt);
            return dt;

        }

        public Dictionary<string,decimal> SelectArea(string db, string khouse)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + db + ";Extended Properties=dBASE IV";
            string strSelect = "SELECT khouse,AREA_1 FROM house_vu ";
            OleDbConnection myConn = new OleDbConnection(strConn);
            OleDbCommand myCommand = new OleDbCommand(strSelect, myConn);
            OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myCommand);
            myConn.Open();
            DataTable dt = new DataTable();
            myDataAdapter.Fill(dt);
            Dictionary<string, decimal> domArea = new Dictionary<string, decimal>();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                if (domArea.ContainsKey(dt.Rows[i][0].ToString()))
                {
                    if (Convert.ToDecimal(dt.Rows[i][1].ToString()) > domArea[dt.Rows[i][0].ToString()])
                    {
                        domArea[dt.Rows[i][0].ToString()] = Convert.ToDecimal(dt.Rows[i][1].ToString());
                    }
                }
                else
                {
                    domArea.Add(dt.Rows[i][0].ToString(), Convert.ToDecimal(dt.Rows[i][1].ToString()));
                }
            }
            return domArea;

        }
    }
}
