using System;
using System.Collections.Generic;
//using System.Linq;
using System.Text;
using System.Data;
using System.Data.OracleClient;


namespace ConsoleApplication9
{
    class Ora
    {
        private string connStr;
        private string connStrTest;
        private string connStrHome;
        public Ora()
        {
            //connStr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.126.128)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=XE)));User Id = HR; Password = test";
            //connStr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.1.30)(PORT=1521))(CONNECT_DATA=(SID=ORCL)));User Id=dbq;Password=dbq;";
            //connStr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.5.77)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=orcl)));User Id = MJF; Password = ActaNonVerba";
            connStr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=85.140.61.250)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ezhkh)));User Id = b4_gkh_samara; Password = ACTANONVERBA";
            connStrTest = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=85.140.61.250)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ezhkh)));User Id = b4_gkh_samara2; Password = ACTANONVERBA";
            connStrHome = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=46.0.13.2)(PORT=1578))(CONNECT_DATA=(SID=orcl)));User Id=dbq;Password=dbq";
        }

        public List<string> houses = new List<string>();

        public DataTable TestHome()
        {
            string cmdText = @"select * FROM LKP_CHARTER";

            OracleConnection conn = new OracleConnection(connStrHome);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            try
            {
                da.Fill(dt);
                return dt;
            }
            catch (Exception e)
            {
                return null;
            }
            finally
            { conn.Close(); }
        }


        public DataTable SelectTableFromDB()
        {
            string cmdText = @"select table_name 
                                from ALL_TABLES 
                                where TABLESPACE_NAME = 'USERS' and owner = 'B4_GKH_SAMARA'
                                order by table_name";

            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            try
            {
                da.Fill(dt);
                return dt;
            }
            catch (Exception e)
            {
                return null;
            }
            finally
            { conn.Close(); }
        }

        public Int32 SelectRowsCount(string dbName, string tableName)
        {
            string cmdText = @"select count(*) FROM " + tableName + " where OBJECT_CREATE_DATE <= to_date('14-02-2016', 'dd-mm-yyyy')";

            OracleConnection conn = new OracleConnection(dbName == "base" ? connStr : connStrTest);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            try
            {
                da.Fill(dt);
                return Convert.ToInt32(dt.Rows[0][0]);
            }
            catch (Exception e)
            {
                return -1;
            }
            finally
            { conn.Close(); }
        }

        public Int32 SelectMaxId(string tableName)
        {
            string cmdText = @"select max(id) FROM " + tableName;

            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            try
            {
                da.Fill(dt);
                return Convert.ToInt32(dt.Rows[0][0]);
            }
            catch (Exception e)
            {
                return -1;
            }
            finally
            { conn.Close(); }
        }

        public string SelectFio(string inn, int position_id)
        {
            string cmdText = @"SELECT full_name FROM gkh_contragent_contact 
where contragent_id = (SELECT id FROM gkh_contragent where inn = " + inn + ") and position_id = " + position_id + " and (date_end_work is null or date_end_work > current_date)";
            //string cmdText = "SELECT id FROM realty_object where mu_name LIKE '%Волжский р-н%'";
            //string cmdText = "SELECT * FROM employees";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            //cmd.Parameters.Add("surname", surname);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            try
            {
                da.Fill(dt);
                string fio = "";
                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        fio += dt.Rows[i][0].ToString() + ",";
                    }
                    fio = fio.Substring(0, fio.Length - 1);
                    return fio;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception e)
            {
                return null;
            }
            finally
            { conn.Close(); }

        }

        public DataTable SelectHouse(string region)
        {
            string cmdText = @"SELECT '3' as code, null as YKDS, gro.gkh_code, gdm.name, gro.address, '666' as ManOrg, t1.category, gro.floors, 
CASE WHEN gro.build_year is not null THEN '01.01.' || gro.build_year ELSE null END as build_year,
area_mkd, null as SMOP, null as sOtopl, SUBSTR(gkh_code, 1, 2) as MO_CODE, null as Dop, number_apartments as rooms, null as PU, t2.kladrcode, 
bfa.place_name as village, 
CASE WHEN bfa.street_name LIKE 'ул. %' THEN SUBSTR(bfa.street_name, 5) 
      WHEN bfa.street_name LIKE 'ш. %' THEN SUBSTR(bfa.street_name, 4) 
      WHEN bfa.street_name LIKE 'мкр. %' THEN SUBSTR(bfa.street_name, 6)
      WHEN bfa.street_name LIKE 'пер. %' THEN SUBSTR(bfa.street_name, 6)
      WHEN bfa.street_name LIKE 'пер. %' THEN SUBSTR(bfa.street_name, 6)
      WHEN bfa.street_name LIKE 'пр-кт. %' THEN SUBSTR(bfa.street_name, 8) ELSE bfa.street_name END as street,
bfa.house as house, bfa.housing as block, gro.address FROM gkh_reality_object gro 
inner join gkh_dict_municipality gdm on gdm.id = gro.municipality_id
inner join b4_fias_address bfa on gro.fias_address_id = bfa.id
left join (SELECT id, CASE WHEN condition_house =20 THEN '7' 
            WHEN otoplenie not in (1,2) OR  HVS not in(1) or Vodotv not in(1) THEN '6'
            WHEN (lift1 = 1 OR lift2 >=1) AND musoroprovod in(1,2) THEN '2'
            WHEN (lift1 = 1 OR lift2 >=1) AND (musoroprovod not in(1,2) or musoroprovod is null) THEN '3'
            WHEN (lift1 = 0 AND (lift2 <1 OR lift2 is null)) AND musoroprovod in(1,2) THEN '4'
            WHEN (lift1 = 0 AND (lift2 <1 OR lift2 is null)) AND (musoroprovod not in(1,2) or musoroprovod is null) THEN '5' END as category FROM(
SELECT gro.id, condition_house, t1.value as otoplenie, t2.value as HVS, t3.value as Vodotv, t4.value as lift1, t41.value as lift2, t5.value as musoroprovod
from gkh_reality_object gro 
LEFT JOIN (SELECT ro.id as id, tpv.value from gkh_reality_object ro
                  inner join tp_teh_passport tp on tp.reality_obj_id = ro.id
                  inner join tp_teh_passport_value tpv on tpv.teh_passport_id = tp.id
                  inner join gkh_dict_municipality gdm on gdm.id = ro.municipality_id
                  where tpv.form_code = 'Form_3_1' AND tpv.cell_code = '1:3') t1 on t1.id = gro.id
LEFT JOIN (SELECT ro.id as id, tpv.value from gkh_reality_object ro
                  inner join tp_teh_passport tp on tp.reality_obj_id = ro.id
                  inner join tp_teh_passport_value tpv on tpv.teh_passport_id = tp.id
                  inner join gkh_dict_municipality gdm on gdm.id = ro.municipality_id
                  where tpv.form_code = 'Form_3_2_CW' AND tpv.cell_code = '1:3') t2 on t2.id = gro.id
LEFT JOIN (SELECT ro.id as id, tpv.value from gkh_reality_object ro
                  inner join tp_teh_passport tp on tp.reality_obj_id = ro.id
                  inner join tp_teh_passport_value tpv on tpv.teh_passport_id = tp.id
                  inner join gkh_dict_municipality gdm on gdm.id = ro.municipality_id
                  where tpv.form_code = 'Form_3_3_Water' AND tpv.cell_code = '1:3') t3 on t3.id = gro.id
LEFT JOIN (SELECT ro.id as id, CASE WHEN ro.floors <9 AND ro.floors is not null THEN 0 ELSE 1 END as value from gkh_reality_object ro) t4 on t4.id = gro.id
LEFT JOIN (SELECT ro.id as id, tpv.value from gkh_reality_object ro
                  inner join tp_teh_passport tp on tp.reality_obj_id = ro.id
                  inner join tp_teh_passport_value tpv on tpv.teh_passport_id = tp.id
                  inner join gkh_dict_municipality gdm on gdm.id = ro.municipality_id
                  where tpv.form_code = 'Form_4_1' AND tpv.cell_code = '1:4' AND tpv.value >0) t41 on t41.id = gro.idнет(=
LEFT JOIN (SELECT ro.id as id, tpv.value from gkh_reality_object ro
                  inner join tp_teh_passport tp on tp.reality_obj_id = ro.id
                  inner join tp_teh_passport_value tpv on tpv.teh_passport_id = tp.id
                  inner join gkh_dict_municipality gdm on gdm.id = ro.municipality_id
                  where tpv.form_code = 'Form_3_7' AND tpv.cell_code = '1:3') t5 on t5.id = gro.id) tab) t1 on t1.id = gro.id
LEFT JOIN (select gro.id, MAX(bf.kladrcode) as kladrcode from gkh_reality_object gro inner join b4_fias_address bfa on bfa.id = gro.fias_address_id 
inner join b4_fias bf on bf.aoguid = bfa.street_guid group by gro.id) t2 on t2.id = gro.id
                            where type_house not in(0, 20) AND condition_house !=40 AND (gkh_code LIKE '87%') AND LENGTH(bfa.house) < 7 AND kladrcode is not null  order by gro.gkh_code";
            //string cmdText = "SELECT id FROM realty_object where mu_name LIKE '%Волжский р-н%'";
            //string cmdText = "SELECT * FROM employees";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            //cmd.Parameters.Add("surname", surname);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            try
            {
                da.Fill(dt);
                return dt;
            }
            catch (Exception e)
            {
                return null;
            }
            finally
            { conn.Close(); }

        }

        public DataTable SelectHouse(List<string> code)
        {
            string gkh_code = "";
            foreach (string cod in code)
            {
                gkh_code += cod + ",";
            }
            gkh_code = gkh_code.Substring(0, gkh_code.Length - 1);
            string cmdText = @"SELECT '3' as code, null as YKDS, gro.gkh_code, gdm.name, gro.address, '666' as ManOrg, t1.category, gro.floors, 
CASE WHEN gro.build_year is not null THEN '01.01.' || gro.build_year ELSE null END as build_year,
area_mkd, null as SMOP, null as sOtopl, SUBSTR(gkh_code, 1, 2) as MO_CODE, null as Dop, number_apartments as rooms, null as PU, t2.kladrcode, 
bfa.place_name as village, 
CASE WHEN bfa.street_name LIKE 'ул. %' THEN SUBSTR(bfa.street_name, 5) 
      WHEN bfa.street_name LIKE 'ш. %' THEN SUBSTR(bfa.street_name, 4) 
      WHEN bfa.street_name LIKE 'мкр. %' THEN SUBSTR(bfa.street_name, 6)
      WHEN bfa.street_name LIKE 'пер. %' THEN SUBSTR(bfa.street_name, 6)
      WHEN bfa.street_name LIKE 'пер. %' THEN SUBSTR(bfa.street_name, 6)
      WHEN bfa.street_name LIKE 'пр-кт. %' THEN SUBSTR(bfa.street_name, 8) ELSE bfa.street_name END as street,
bfa.house as house, bfa.housing as block, gro.address FROM gkh_reality_object gro 
inner join gkh_dict_municipality gdm on gdm.id = gro.municipality_id
inner join b4_fias_address bfa on gro.fias_address_id = bfa.id
left join (SELECT id, CASE WHEN condition_house =20 THEN '7' 
            WHEN otoplenie not in (1,2) OR  HVS not in(1) or Vodotv not in(1) THEN '6'
            WHEN (lift1 = 1 OR lift2 >=1) AND musoroprovod in(1,2) THEN '2'
            WHEN (lift1 = 1 OR lift2 >=1) AND (musoroprovod not in(1,2) or musoroprovod is null) THEN '3'
            WHEN (lift1 = 0 AND (lift2 <1 OR lift2 is null)) AND musoroprovod in(1,2) THEN '4'
            WHEN (lift1 = 0 AND (lift2 <1 OR lift2 is null)) AND (musoroprovod not in(1,2) or musoroprovod is null) THEN '5' END as category FROM(
SELECT gro.id, condition_house, t1.value as otoplenie, t2.value as HVS, t3.value as Vodotv, t4.value as lift1, t41.value as lift2, t5.value as musoroprovod
from gkh_reality_object gro 
LEFT JOIN (SELECT ro.id as id, tpv.value from gkh_reality_object ro
                  inner join tp_teh_passport tp on tp.reality_obj_id = ro.id
                  inner join tp_teh_passport_value tpv on tpv.teh_passport_id = tp.id
                  inner join gkh_dict_municipality gdm on gdm.id = ro.municipality_id
                  where tpv.form_code = 'Form_3_1' AND tpv.cell_code = '1:3') t1 on t1.id = gro.id
LEFT JOIN (SELECT ro.id as id, tpv.value from gkh_reality_object ro
                  inner join tp_teh_passport tp on tp.reality_obj_id = ro.id
                  inner join tp_teh_passport_value tpv on tpv.teh_passport_id = tp.id
                  inner join gkh_dict_municipality gdm on gdm.id = ro.municipality_id
                  where tpv.form_code = 'Form_3_2_CW' AND tpv.cell_code = '1:3') t2 on t2.id = gro.id
LEFT JOIN (SELECT ro.id as id, tpv.value from gkh_reality_object ro
                  inner join tp_teh_passport tp on tp.reality_obj_id = ro.id
                  inner join tp_teh_passport_value tpv on tpv.teh_passport_id = tp.id
                  inner join gkh_dict_municipality gdm on gdm.id = ro.municipality_id
                  where tpv.form_code = 'Form_3_3_Water' AND tpv.cell_code = '1:3') t3 on t3.id = gro.id
LEFT JOIN (SELECT ro.id as id, CASE WHEN ro.floors <9 AND ro.floors is not null THEN 0 ELSE 1 END as value from gkh_reality_object ro) t4 on t4.id = gro.id
LEFT JOIN (SELECT ro.id as id, tpv.value from gkh_reality_object ro
                  inner join tp_teh_passport tp on tp.reality_obj_id = ro.id
                  inner join tp_teh_passport_value tpv on tpv.teh_passport_id = tp.id
                  inner join gkh_dict_municipality gdm on gdm.id = ro.municipality_id
                  where tpv.form_code = 'Form_4_1' AND tpv.cell_code = '1:4' AND tpv.value >0) t41 on t41.id = gro.id
LEFT JOIN (SELECT ro.id as id, tpv.value from gkh_reality_object ro
                  inner join tp_teh_passport tp on tp.reality_obj_id = ro.id
                  inner join tp_teh_passport_value tpv on tpv.teh_passport_id = tp.id
                  inner join gkh_dict_municipality gdm on gdm.id = ro.municipality_id
                  where tpv.form_code = 'Form_3_7' AND tpv.cell_code = '1:3') t5 on t5.id = gro.id) tab) t1 on t1.id = gro.id
LEFT JOIN (select gro.id, MAX(bf.kladrcode) as kladrcode from gkh_reality_object gro inner join b4_fias_address bfa on bfa.id = gro.fias_address_id 
inner join b4_fias bf on bf.aoguid = bfa.street_guid group by gro.id) t2 on t2.id = gro.id
                            where type_house not in(0, 20) AND condition_house !=40 AND (gkh_code in("+gkh_code+")) AND LENGTH(bfa.house) < 7 AND kladrcode is not null  order by gro.gkh_code";
            //string cmdText = "SELECT id FROM realty_object where mu_name LIKE '%Волжский р-н%'";
            //string cmdText = "SELECT * FROM employees";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            //cmd.Parameters.Add("surname", surname);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            try
            {
                da.Fill(dt);
                return dt;
            }
            catch (Exception e)
            {
                return null;
            }
            finally
            { conn.Close(); }

        }

        public DataTable SelectLN(string region)
        {
            string cmdText = @"SELECT '4' as code, substr(gsa.owner_code,1,7) as house_code, gsa.owner_code,  '1' as typeLS, goai.fio_owner, goai.num_apartment, goai.count_people, '0' as COunt17, '0' as count18, '1' as count19, goai.area_total,
 goai.area_living, '0' as count24, CASE WHEN goai.privatized = 20 THEN 0 ELSE 1 END as qwerty
 FROM gkh_obj_apartment_info goai
inner join gkh_sam_ap_inf_own_code gsa on goai.id = gsa.ap_info_id
 where reality_object_id in (
 SELECT gro.id from gkh_reality_object gro
 LEFT JOIN (select gro.id, MAX(bf.kladrcode) as kladrcode from gkh_reality_object gro inner join b4_fias_address bfa on bfa.id = gro.fias_address_id 
inner join b4_fias bf on bf.aoguid = bfa.street_guid where LENGTH(bfa.house) < 7 group by gro.id) t2 on t2.id = gro.id 
where (gkh_code LIKE '97%' OR gkh_code LIKE '98%' OR gkh_code LIKE '99%') AND kladrcode is not null AND type_house not in(0, 20) AND condition_house !=40)
order by house_code desc";
            //string cmdText = "SELECT id FROM realty_object where mu_name LIKE '%Волжский р-н%'";
            //string cmdText = "SELECT * FROM employees";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            //cmd.Parameters.Add("surname", surname);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            try
            {
                da.Fill(dt);
                return dt;
            }
            catch (Exception e)
            {
                return null;
            }
            finally
            { conn.Close(); }

        }

        public DataTable SelectLN2(string region)
        {
            string cmdText = @"SELECT gkh_code,substr(gkh_code, 1,2) as code_reg, gro.id, goai.num_apartment, goai.fio_owner, goai.area_total 
                            from gkh_reality_object gro 
                            inner join gkh_obj_apartment_info goai on goai.reality_object_id = gro.id 
                            where type_house not in(0, 20) AND condition_house !=40 AND substr(gkh_code, 1,2) not in(88,89,90,91,92,93,94,95,96,97,98,99)  order by gkh_code";
            //string cmdText = "SELECT id FROM realty_object where mu_name LIKE '%Волжский р-н%'";
            //string cmdText = "SELECT * FROM employees";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            //cmd.Parameters.Add("surname", surname);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            try
            {
                da.Fill(dt);
                return dt;
            }
            catch (Exception e)
            {
                return null;
            }
            finally
            { conn.Close(); }

        }

        public DataTable SelectLN3(string region)
        {
            string cmdText = @"SELECT '4' as code, gro.gkh_code, '2' as typeLS, gohi.name, gohi.owner, '-' as num_appartment,
'0' as countPeople, '0' as COunt17, '0' as count18, '1' as count19, gohi.total_area, null as area_living, '0' as count24
FROM gkh_obj_house_info gohi
INNER JOIN gkh_reality_object gro on gro.id = gohi.reality_object_id
 where reality_object_id in (
 SELECT gro.id from gkh_reality_object gro
 LEFT JOIN (select gro.id, MAX(bf.kladrcode) as kladrcode from gkh_reality_object gro inner join b4_fias_address bfa on bfa.id = gro.fias_address_id 
inner join b4_fias bf on bf.aoguid = bfa.street_guid where LENGTH(bfa.house) < 7 group by gro.id) t2 on t2.id = gro.id 
where (gkh_code LIKE '" + region + "%') AND kladrcode is not null AND type_house not in(0, 20) AND condition_house !=40) AND gkh_code not in(8700479, 8700638, 8201811)";
            //string cmdText = "SELECT id FROM realty_object where mu_name LIKE '%Волжский р-н%'";
            //string cmdText = "SELECT * FROM employees";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            //cmd.Parameters.Add("surname", surname);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            try
            {
                da.Fill(dt);
                return dt;
            }
            catch (Exception e)
            {
                return null;
            }
            finally
            { conn.Close(); }

        }

        public DataTable SelectLN4(string region)
        {
            string cmdText = @"SELECT '4' as code, gro.gkh_code as house_code,0 as owner_code,  '1' as typeLS, goai.fio_owner, goai.num_apartment, goai.count_people, '0' as COunt17, '0' as count18, '1' as count19, goai.area_total,
 goai.area_living, '0' as count24, CASE WHEN goai.privatized = 20 THEN 0 ELSE 1 END as qwerty
 FROM gkh_obj_apartment_info goai
 inner join gkh_reality_object gro on gro.id = goai.reality_object_id
 where reality_object_id in (
 SELECT gro.id from gkh_reality_object gro
 LEFT JOIN (select gro.id, MAX(bf.kladrcode) as kladrcode from gkh_reality_object gro inner join b4_fias_address bfa on bfa.id = gro.fias_address_id 
inner join b4_fias bf on bf.aoguid = bfa.street_guid where LENGTH(bfa.house) < 7 group by gro.id) t2 on t2.id = gro.id 
where (gkh_code LIKE '76%') AND kladrcode is not null AND type_house not in(0, 20) AND condition_house !=40 AND gkh_code in(7600125, 7600041, 7600039, 7600010, 7600040, 7600044, 7600029)) 
order by house_code desc";
            //string cmdText = "SELECT id FROM realty_object where mu_name LIKE '%Волжский р-н%'";
            //string cmdText = "SELECT * FROM employees";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            //cmd.Parameters.Add("surname", surname);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            try
            {
                da.Fill(dt);
                return dt;
            }
            catch (Exception e)
            {
                return null;
            }
            finally
            { conn.Close(); }

        }

        public DataTable SelectLN4(List<string> code)
        {
            string gkh_code = "";
            foreach (string cod in code)
            {
                gkh_code += cod + ",";
            }
            gkh_code = gkh_code.Substring(0, gkh_code.Length - 1);
            string cmdText = @"SELECT '4' as code, gro.gkh_code as house_code,0 as owner_code,  '1' as typeLS, goai.fio_owner, goai.num_apartment, goai.count_people, '0' as COunt17, '0' as count18, '1' as count19, goai.area_total,
 goai.area_living, '0' as count24, CASE WHEN goai.privatized = 20 THEN 0 ELSE 1 END as qwerty
 FROM gkh_obj_apartment_info goai
 inner join gkh_reality_object gro on gro.id = goai.reality_object_id
 where reality_object_id in (
 SELECT gro.id from gkh_reality_object gro
 LEFT JOIN (select gro.id, MAX(bf.kladrcode) as kladrcode from gkh_reality_object gro inner join b4_fias_address bfa on bfa.id = gro.fias_address_id 
inner join b4_fias bf on bf.aoguid = bfa.street_guid where LENGTH(bfa.house) < 7 group by gro.id) t2 on t2.id = gro.id 
where  kladrcode is not null AND type_house not in(0, 20) AND condition_house !=40 AND gkh_code in(" + gkh_code + ")) order by house_code,  goai.num_apartment";
            //string cmdText = "SELECT id FROM realty_object where mu_name LIKE '%Волжский р-н%'";
            //string cmdText = "SELECT * FROM employees";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            //cmd.Parameters.Add("surname", surname);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            try
            {
                da.Fill(dt);
                return dt;
            }
            catch (Exception e)
            {
                return null;
            }
            finally
            { conn.Close(); }

        }

        public DataTable SelectGkhCode(string ulica, string dom)
        {
            string cmdText = "SELECT gro.gkh_code, gro.address from GKH_REALITY_OBJECT gro where replace(lower(GRO.ADDRESS), ' ','') LIKE replace(lower('%" + ulica + "%'), ' ','') " +
                //" and MUNICIPALITY_ID in (21690, 21691, 21692, 21693, 21694, 21695, 21696, 21697, 21698) and replace(lower(GRO.ADDRESS), ' ','') LIKE replace(lower('%д. " + dom + "'), ' ','')";
                " and MUNICIPALITY_ID in (21684) and replace(lower(GRO.ADDRESS), ' ','') LIKE replace(lower('%д. " + dom + "'), ' ','')";
            //string cmdText = "SELECT id FROM realty_object where mu_name LIKE '%Волжский р-н%'";
            //string cmdText = "SELECT * FROM employees";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            //cmd.Parameters.Add("surname", surname);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
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

        public Int32 SelectGkhCode(string address)
        {
            string cmdText = "SELECT gro.gkh_code from GKH_REALITY_OBJECT gro where replace(lower(GRO.ADDRESS), ' ','') = replace(lower('%" + address + "%'), ' ','') " +
                " and MUNICIPALITY_ID in (21690, 21691, 21692, 21693, 21694, 21695, 21696, 21697, 21698)";
            //string cmdText = "SELECT id FROM realty_object where mu_name LIKE '%Волжский р-н%'";
            //string cmdText = "SELECT * FROM employees";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            //cmd.Parameters.Add("surname", surname);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
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

        public DataTable SelectDataToRep(Int32 gkhCode)
        {
            string cmdText = @"SELECT to_char(gro.DATE_COMMISSIONING, 'dd.mm.yyyy') as DATE_COMMISSIONING, gro.AREA_MKD, gro.NUMBER_APARTMENTS, t1.val as NUMBER_APARTMENTS_NOT_LIVING, 
                                gro.AREA_LIV_NOT_LIV_MKD, gro.area_living,  gro.AREA_LIV_NOT_LIV_MKD - gro.area_living as area_not_living, gro.number_living
                                from GKH_REALITY_OBJECT gro 
                                LEFT JOIN (SELECT ttpv.value as val, ttp.reality_obj_id FROM TP_TEH_PASSPORT ttp INNER JOIN TP_TEH_PASSPORT_VALUE ttpv on ttp.id = ttpv.TEH_PASSPORT_ID 
                                           where form_code = 'Form_1_3' and cell_code = '1:1') t1 on t1.reality_obj_id = gro.id
                                where GKH_CODE = " + gkhCode;
            //string cmdText = "SELECT id FROM realty_object where mu_name LIKE '%Волжский р-н%'";
            //string cmdText = "SELECT * FROM employees";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            //cmd.Parameters.Add("surname", surname);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
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

        public string UpdatePeople(string gkh_code, Int32 flat, string fio,
            string total_area, string owner_area, string priv, string people_count)
        {
            //string cmdText = "SELECT gro.id from GKH_REALITY_OBJECT gro inner join GKH_DICT_MUNICIPALITY gdm on GDM.ID = GRO.MUNICIPALITY_ID where GDM.NAME LIKE '%Жигулевск' AND replace(lower(GRO.ADDRESS), ' ','') = replace(lower('" + gkh_code + "'), ' ','')";
            string cmdText = "SELECT id from GKH_REALITY_OBJECT where gkh_code = " + gkh_code;
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            int id;
            conn.Open();
            try
            {
                da.Fill(dt);
                id = Convert.ToInt32(dt.Rows[0][0]);
                string cmdCheck = "SELECT id, privatized from gkh_obj_apartment_info where reality_object_id = " + id +
                    " and num_apartment = '" + flat + "'";
                OracleCommand cmd2 = new OracleCommand(cmdCheck, conn);
                OracleDataAdapter da2 = new OracleDataAdapter(cmd2);
                DataTable dt2 = new DataTable();
                int id2;
                int privatized;
                da2.Fill(dt2);
                if (dt2.Rows.Count == 1)
                {
                    id2 = Convert.ToInt32(dt2.Rows[0][0]);
                    privatized = Convert.ToInt32(dt2.Rows[0][1]);
                    if (total_area == null || total_area == "" || total_area == " ")
                        total_area = "0";
                    int privat = 30;
                    if (priv == "да" || priv == "Да")
                        privat = 10;
                    else
                        privat = 20;
                    if (people_count == null || people_count == "" || people_count == " ")
                        people_count = "0";
                    string cmdText1 = "UPDATE gkh_obj_apartment_info set area_total = " + total_area + ", " +
                            "fio_owner = '" + fio + "', count_people = " + people_count;
                    if (privatized != 990)
                        cmdText1 += ", privatized = " + privat;
                    cmdText1 += " where id = " + id2;
                    OracleCommand cmd1 = new OracleCommand(cmdText1, conn);

                    cmd1.ExecuteNonQuery();
                    return "ЗАГРУЖЕНО";
                }
                else if (dt2.Rows.Count == 0)
                {
                    if (total_area == null || total_area == "" || total_area == " ")
                        total_area = "0";
                    int privat = 30;
                    if (priv == "да" || priv == "Да")
                        privat = 10;
                    else if (priv == " ")
                        privat = 30;
                    else
                        privat = 20;
                    if (people_count == null || people_count == "" || people_count == " ")
                        people_count = "0";
                    string cmdText1 = "INSERT INTO gkh_obj_apartment_info (id, object_version, object_create_date, object_edit_date, num_apartment, area_total, privatized," +
                    "reality_object_id, fio_owner, count_people) VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, '" + flat + "'," + total_area
                    + "," + privat + "," + id + ",'" + fio + "', " + people_count + ")";
                    OracleCommand cmd1 = new OracleCommand(cmdText1, conn);
                    cmd1.ExecuteNonQuery();
                    return "ЗАГРУЖЕНО";
                }
                else
                {
                    return "БОЛЕЕ 1-ой КВАРТИРЫ - " + flat;
                }
            }
            catch (Exception e)
            {
                string err = e.Message;
                return err + "|" + flat;
            }
            finally
            { conn.Close(); }

        }

        public string InsertPeople(string gkh_code, string flat, string fio)
        {
            string cmdText = "SELECT id from GKH_REALITY_OBJECT where gkh_code = " + gkh_code;
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
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
                OracleCommand cmd2 = new OracleCommand(cmdCheck, conn);
                OracleDataAdapter da2 = new OracleDataAdapter(cmd2);
                DataTable dt2 = new DataTable();
                int id2;
                da2.Fill(dt2);
                id2 = Convert.ToInt32(dt2.Rows[0][0]);
                string cmdText1 = "UPDATE gkh_obj_apartment_info set " +
                        "fio_owner = '" + fio + "' where id = " + id2;
                OracleCommand cmd1 = new OracleCommand(cmdText1, conn);

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

        public string InsertPeople(string gkh_code, string flat, string fio, string total_area, string people_count)
        {
            string cmdText = "SELECT id from GKH_REALITY_OBJECT where gkh_code = " + gkh_code;
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
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
                OracleCommand cmd2 = new OracleCommand(cmdCheck, conn);
                OracleDataAdapter da2 = new OracleDataAdapter(cmd2);
                DataTable dt2 = new DataTable();
                int id2;
                da2.Fill(dt2);
                id2 = Convert.ToInt32(dt2.Rows[0][0]);
                if (flat.Substring(0, 1) == "0")
                    flat = flat.Substring(1);
                if (flat.Substring(0, 1) == "0")
                    flat = flat.Substring(1);
                if (total_area == null || total_area == "" || total_area == " ")
                    total_area = "0";
                if (people_count == null || people_count == "" || people_count == " ")
                    people_count = "0";
                string cmdText1 = "UPDATE gkh_obj_apartment_info set " +
                        "fio_owner = '" + fio + "', area_total = " + total_area + ", count_people = " + people_count + " where id = " + id2;
                OracleCommand cmd1 = new OracleCommand(cmdText1, conn);

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

        public string InsertOffice(string gkh_code, string flat, string fio, string area)
        {
            string cmdText = "SELECT id from GKH_REALITY_OBJECT where gkh_code = " + gkh_code;
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            int id;
            conn.Open();
            da.Fill(dt);
            id = Convert.ToInt32(dt.Rows[0][0]);
            string total_area = "0";
            if (area != "не начисл." && area != "")
                total_area = area.Replace(',', '.');
            string cmdText1 = "INSERT INTO gkh_obj_house_info(id, object_version, object_create_date, object_edit_date," +
                    "reality_object_id, num, name, owner, total_area) VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, '" + id + "','" + flat
                    + "','Офис','" + fio + "'," + total_area + ")";
            OracleCommand cmd1 = new OracleCommand(cmdText1, conn);
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

        public string InsertOffice(string gkh_code, string name, string flat, string area, string vidPrava, string numPrava, string dateReg, string owner, string Birtday)
        {
            //string cmdText = "SELECT id from GKH_REALITY_OBJECT where gkh_code = " + gkh_code;
            string cmdText = "SELECT gro.id from GKH_REALITY_OBJECT gro inner join GKH_DICT_MUNICIPALITY gdm on GDM.ID = GRO.MUNICIPALITY_ID where GDM.NAME LIKE '%г.о. Жигулевск' AND replace(lower(GRO.ADDRESS), ' ','') = replace(lower('" + gkh_code + "'), ' ','')";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            int id;
            conn.Open();
            da.Fill(dt);
            id = Convert.ToInt32(dt.Rows[0][0]);
            if (!houses.Contains(id.ToString()))
            {
                houses.Add(id.ToString());
                string cmdText2 = "DELETE FROM gkh_obj_house_info where name = 'Жилое помещение' AND reality_object_id = " + id;
                OracleCommand cmd2 = new OracleCommand(cmdText2, conn);
                cmd2.ExecuteNonQuery();
            }
            int kind_right = 0;
            if (vidPrava != null && vidPrava != "")
                kind_right = 30;
            string cmdText1 = "INSERT INTO gkh_obj_house_info(id, object_version, object_create_date, object_edit_date," +
                    "reality_object_id, num, name, owner, total_area, kind_right, num_reg_right, date_reg, date_reg_owner) VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, '" + id + "','" + flat
                    + "','" + name + "','" + owner + "'," + area + ", " + kind_right + ", '" + numPrava + "', to_date('" + dateReg.Substring(0,10) 
                    + "', 'dd.mm.yyyy'), to_date('" + Birtday + "','dd.mm.yyyy'))";
            OracleCommand cmd1 = new OracleCommand(cmdText1, conn);
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

        public string InsertHouseManOrg(string date, string manOrgId)
        {
            string cmdText = "select * from gkh_man_org_real_obj where manag_org_id =" + manOrgId + " AND reality_object_id not in (SELECT reality_obj_id FROM gkh_morg_contract_realobj )";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            int id;
            conn.Open();
            da.Fill(dt);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string cmdText1 = "INSERT INTO gkh_morg_contract (ID, object_version, object_create_date, object_edit_date, manag_org_id, type_contract, start_date" +
                        ") VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, '" + manOrgId + "', 20" +
                        ",TO_TIMESTAMP('" + date + "', 'DD.MM.YYYY')) RETURNING id INTO :id";
                OracleCommand cmd1 = new OracleCommand(cmdText1, conn);
                cmd1.Parameters.Add(new OracleParameter
                {
                    ParameterName = ":id",
                    OracleType = OracleType.Number,
                    Direction = ParameterDirection.Output
                });
                try
                {
                    cmd1.ExecuteNonQuery();
                    int id2 = Convert.ToInt32(cmd1.Parameters[":id"].Value.ToString());
                    string cmdText2 = "INSERT INTO gkh_morg_contract_realobj (id, object_version, object_create_date, object_edit_date, reality_obj_id, man_org_contract_id" +
                        ") VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, '" + dt.Rows[i][5].ToString() +
                        "','" + id2 + "')";
                    OracleCommand cmd2 = new OracleCommand(cmdText2, conn);
                    cmd2.ExecuteNonQuery();

                    string cmdText3 = "INSERT INTO gkh_morg_contract_owners (id, object_version, object_create_date, object_edit_date, contract_foundation" +
                        ") VALUES(" + id2 + ", 0, CURRENT_DATE, CURRENT_DATE, 20)";
                    OracleCommand cmd3 = new OracleCommand(cmdText3, conn);
                    cmd3.ExecuteNonQuery();
                }
                catch (Exception e)
                {
                    string err = e.Message;
                    conn.Close(); 
                    return err;
                }
            }
            conn.Close(); 
            return "ЗАГРУЖЕНО";
        }

        public string InsertCommunalOrg2(string manOrgId, string gkh_code)
        {
            string cmdText = "SELECT id from GKH_REALITY_OBJECT where gkh_code = " + gkh_code;
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            int id;
            conn.Open();
            da.Fill(dt);
            id = Convert.ToInt32(dt.Rows[0][0]);
            string cmdText1 = "INSERT INTO gkh_supply_resorg_ro (ID, object_version, object_create_date, object_edit_date, supply_resorg_id, realityobject_id" +
                      ") VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, '" + manOrgId + "', "+id+")RETURNING id INTO :id";
            OracleCommand cmd1 = new OracleCommand(cmdText1, conn);
            cmd1.Parameters.Add(new OracleParameter
            {
                ParameterName = ":id",
                OracleType = OracleType.Number,
                Direction = ParameterDirection.Output
            });
            try
            {
                cmd1.ExecuteNonQuery();
                int id2 = Convert.ToInt32(cmd1.Parameters[":id"].Value.ToString());
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    cmdText1 = "INSERT INTO gkh_obj_resorg (ID, object_version, object_create_date, object_edit_date, resorg_id, reality_object_id, date_start" +
                        ") VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, '" + manOrgId + "'," + id + "," +
                        "TO_TIMESTAMP('01.01.2011', 'DD.MM.YYYY'))";
                    OracleCommand cmd2 = new OracleCommand(cmdText1, conn);
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

        public string InsertPeople2(string gkh_code, Int32 flat, string fio,
            string total_area, string useful_area, string privatized, string residents_count)
        {
            //string cmdText = "SELECT gro.id from GKH_REALITY_OBJECT gro inner join GKH_DICT_MUNICIPALITY gdm on GDM.ID = GRO.MUNICIPALITY_ID where GDM.NAME LIKE '%г. Самара, Красноглинский р-н' AND replace(lower(GRO.ADDRESS), ' ','') = replace(lower('" + gkh_code + "'), ' ','')";
            string cmdText = "SELECT id from GKH_REALITY_OBJECT where gkh_code = " + gkh_code;
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            int id;
            try
            {
                conn.Open();
            }
            catch (Exception ex)
            {
                {}
                throw;
            }
            
            da.Fill(dt);
            id = Convert.ToInt32(dt.Rows[0][0]);
            int priv = 30;
            if (privatized.ToLower() == "да" || privatized.ToLower() == "Да")
                priv = 10;
            else
                priv = 20; 
            if (useful_area == null || useful_area == "" || useful_area == " ")
                useful_area = "0";
            if (total_area == null || total_area == "" || total_area == " ")
                total_area = "0";
            if (residents_count == null || residents_count == "" || residents_count == " ")
                residents_count = "0";
            useful_area = useful_area.Replace(',', '.');
            total_area = total_area.Replace(',', '.');
            if (!houses.Contains(id.ToString()))
            {
                houses.Add(id.ToString());
                string cmdText2 = "DELETE FROM gkh_obj_apartment_info where reality_object_id = " + id;
                OracleCommand cmd2 = new OracleCommand(cmdText2, conn);
                cmd2.ExecuteNonQuery();
            }

            string cmdText1 = "INSERT INTO gkh_obj_apartment_info (id, object_version, object_create_date, object_edit_date, num_apartment, area_total, count_people, privatized," +
                    "reality_object_id, fio_owner, area_living) VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, '" + flat + "'," + total_area
                    + "," + residents_count + "," + priv + "," + id + ",'" + fio + "'," + useful_area + ")";
            OracleCommand cmd1 = new OracleCommand(cmdText1, conn);
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

        public string InsertPeople3(string gkh_code, string flat, string fio,
            string useful_area, string total_area, string residents_count, string privatized)
        {
            //string cmdText = "SELECT gro.id from GKH_REALITY_OBJECT gro inner join GKH_DICT_MUNICIPALITY gdm on GDM.ID = GRO.MUNICIPALITY_ID where GDM.NAME LIKE '%г. Самара, Красноглинский р-н' AND replace(lower(GRO.ADDRESS), ' ','') = replace(lower('" + gkh_code + "'), ' ','')";
            string cmdText = "SELECT id from GKH_REALITY_OBJECT where gkh_code = " + gkh_code;
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            int id;
            conn.Open();
            da.Fill(dt);
            id = Convert.ToInt32(dt.Rows[0][0]);
            int priv = 30;
            if (privatized == "да" || privatized == "Да")
                priv = 10;
            else
                priv = 20;
            if (useful_area == null || useful_area == "" || useful_area == " ")
                useful_area = "0";
            if (total_area == null || total_area == "" || total_area == " ")
                total_area = "0";
            if (residents_count == null || residents_count == "" || residents_count == " ")
                residents_count = "0";

            if (!houses.Contains(id.ToString()))
            {
                houses.Add(id.ToString());
                string cmdText2 = "DELETE FROM gkh_obj_apartment_info where reality_object_id = " + id;
                OracleCommand cmd2 = new OracleCommand(cmdText2, conn);
                cmd2.ExecuteNonQuery();
            }

            string cmdText1 = "INSERT INTO gkh_obj_apartment_info (id, object_version, object_create_date, object_edit_date, num_apartment, area_total, count_people, privatized," +
                    "reality_object_id, fio_owner, area_living) VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, '" + flat + "'," + total_area
                    + "," + residents_count + "," + priv + "," + id + ",'" + fio + "'," + useful_area + ")";
            OracleCommand cmd1 = new OracleCommand(cmdText1, conn);
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

        public string InsertPeople3(string gkh_code, string flat, string fio,
            string total_area, string privatized, string residents_count)
        {
            //string cmdText = "SELECT gro.id from GKH_REALITY_OBJECT gro inner join GKH_DICT_MUNICIPALITY gdm on GDM.ID = GRO.MUNICIPALITY_ID where GDM.NAME LIKE '%г. Самара, Красноглинский р-н' AND replace(lower(GRO.ADDRESS), ' ','') = replace(lower('" + gkh_code + "'), ' ','')";
            string cmdText = "SELECT id from GKH_REALITY_OBJECT where gkh_code = " + gkh_code;
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
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
            if (total_area == null || total_area == "" || total_area == " ")
                total_area = "0";
            if (residents_count == null || residents_count == "" || residents_count == " ")
                residents_count = "0";

            if (!houses.Contains(id.ToString()))
            {
                houses.Add(id.ToString());
                string cmdText2 = "DELETE FROM gkh_obj_apartment_info where reality_object_id = " + id;
                OracleCommand cmd2 = new OracleCommand(cmdText2, conn);
                cmd2.ExecuteNonQuery();
            }

            string cmdText1 = "INSERT INTO gkh_obj_apartment_info (id, object_version, object_create_date, object_edit_date, num_apartment, area_total, count_people, privatized," +
                    "reality_object_id, fio_owner) VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, '" + flat + "'," + total_area.Replace(',', '.')
                    + "," + residents_count + "," + priv + "," + id + ",'" + fio + "')";
            OracleCommand cmd1 = new OracleCommand(cmdText1, conn);
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

        public string InsertPeople5(string gkh_code, Int32 flat, string total_area, string useful_area, string privatized, string residents_count, string fio)
        {
            //string cmdText = "SELECT gro.id from GKH_REALITY_OBJECT gro inner join GKH_DICT_MUNICIPALITY gdm on GDM.ID = GRO.MUNICIPALITY_ID where GDM.NAME LIKE '%г. Самара, Красноглинский р-н' AND replace(lower(GRO.ADDRESS), ' ','') = replace(lower('" + gkh_code + "'), ' ','')";
            string cmdText = "SELECT id from GKH_REALITY_OBJECT where gkh_code = " + gkh_code;
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
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
            if (total_area == null || total_area == "" || total_area == " ")
                total_area = "0";
            if (useful_area == null || useful_area == "" || useful_area == " ")
                useful_area = "0";
            if (residents_count == null || residents_count == "" || residents_count == " ")
                residents_count = "0";

            if (!houses.Contains(id.ToString()))
            {
                houses.Add(id.ToString());
                string cmdText2 = "DELETE FROM gkh_obj_apartment_info where reality_object_id = " + id;
                OracleCommand cmd2 = new OracleCommand(cmdText2, conn);
                cmd2.ExecuteNonQuery();
            }

            string cmdText1 = "INSERT INTO gkh_obj_apartment_info (id, object_version, object_create_date, object_edit_date, num_apartment, area_total, count_people, privatized," +
                    "reality_object_id, fio_owner, area_living) VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, '" + flat + "'," + total_area.Replace(',', '.')
                    + "," + residents_count + "," + priv + "," + id + ",'" + fio + "', " + useful_area.Replace(',', '.') + ")";
            OracleCommand cmd1 = new OracleCommand(cmdText1, conn);
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


        public string InsertPeople2(string gkh_code, string flat, string fio, string total_area, string people_count)
        {
            string cmdText = "SELECT id from GKH_REALITY_OBJECT where gkh_code = " + gkh_code;
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            int id;

            string cmdText2 = "SELECT max(id) from gkh_obj_apartment_info";
            OracleCommand cmd2 = new OracleCommand(cmdText2, conn);
            OracleDataAdapter da2 = new OracleDataAdapter(cmd2);
            DataTable dt2 = new DataTable();
            int id2;
            conn.Open();
            try
            {
                da.Fill(dt);
                id = Convert.ToInt32(dt.Rows[0][0]);

                da2.Fill(dt2);
                id2 = Convert.ToInt32(dt2.Rows[0][0]) + 1;
                if (flat.Substring(0, 1) == "0")
                    flat = flat.Substring(1);
                if (flat.Substring(0, 1) == "0")
                    flat = flat.Substring(1);
                int priv = 30;
                if (total_area == null || total_area == "" || total_area == " ")
                    total_area = "0";
                if (people_count == null || people_count == "" || people_count == " ")
                    people_count = "0";

                string cmdText1 = "INSERT INTO gkh_obj_apartment_info (id, object_version, object_create_date, object_edit_date, num_apartment, area_total, count_people, privatized," +
                   "reality_object_id, fio_owner, area_living) VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, '" + flat + "'," + total_area
                   + "," + people_count + "," + priv + "," + id + ",'" + fio + "', 0" + ")";
                OracleCommand cmd1 = new OracleCommand(cmdText1, conn);

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

        public DataTable SelectDubl()
        {
            string cmdText = @"SELECT meta_attribute_id, house_prov_passport_id, group_key, count(*) 
FROM gkh_house_prov_pass_row
group by meta_attribute_id, house_prov_passport_id, group_key
having count(*)>1";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            try
            {
                da.Fill(dt);
                return dt;
            }
            catch (Exception e)
            {
                return null;
            }
            finally
            { conn.Close(); }

        }

        public DataTable SelectKoap()
        {
            string cmdText = @"SELECT gd.DOCUMENT_NUMBER, gda.name
FROM gji_document gd
LEFT JOIN GJI_PROTOCOL_ARTLAW gpa on gpa.protocol_id = gd.id
LEFT JOIN GJI_DICT_ARTICLELAW gda on gda.id = gpa.ARTICLELAW_ID
where gd.type_document = 60 AND gd.document_date >= to_date('01-01-2015', 'dd-mm-yyyy') AND gd.document_date < to_date('31-12-2015', 'dd-mm-yyyy') 
and gda.name in ('ч.1. ст 14.1.3 КоАП РФ', 'ч.2 ст. 14.1.3 КоАП РФ', 'ч.1 ст. 14.1.3 КоАП РФ', 'ч.24 ст 19.5 КоАП РФ', 'ч.2 ст. 7.23.3 КоАП РФ', 'ч.1 ст. 7.23.3 КоАП РФ')
order by gda.name";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            try
            {
                da.Fill(dt);
                return dt;
            }
            catch (Exception e)
            {
                return null;
            }
            finally
            { conn.Close(); }

        }

        public void DelDubl(DataRow dr)
        {
            string cmdText = @"delete from gkh_house_prov_pass_row where house_prov_passport_id = "+ dr[1].ToString()+" and meta_attribute_id = " + dr[0].ToString() +
                " and group_key = " + dr[2].ToString() + " and id not in (select min(id) FROM gkh_house_prov_pass_row where house_prov_passport_id = " + dr[1].ToString() +
                " and meta_attribute_id = " + dr[0].ToString() + " and group_key = " + dr[2].ToString() + ")";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd1 = new OracleCommand(cmdText, conn);
            
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

        public DataTable SelectCurRepair(string inn, string year)
        {
            string cmdText = @"SELECT gro.gkh_code, gro.address, gdwcr.name, plan_date, plan_sum, plan_work, fact_date, fact_sum, fact_work, inn 
FROM gkh_obj_curent_repair gocr
INNER JOIN gkh_reality_object gro on gro.id = gocr.reality_object_id
INNER JOIN(SELECT t3.reality_object_id, gc.id, gc.name as name, gdm.name || ', ' || gc.juridical_address as juridical_address, gc.inn
                      FROM(select t2.reality_object_id, mo.contragent_id from(SELECT manag_org_id, gmcr.reality_obj_id as reality_object_id, max(gmcr.id) 
                  FROM gkh_morg_contract gmc
                  INNER JOIN gkh_morg_contract_realobj gmcr on gmcr.man_org_contract_id = gmc.id
                  where current_date>start_date AND(current_date < end_date OR end_date is null)
                  group by manag_org_id, gmcr.reality_obj_id) t2
                              left join gkh_managing_organization mo on mo.id = t2.manag_org_id) t3 inner join gkh_contragent gc on gc.id = t3.contragent_id
                              inner join gkh_dict_municipality gdm on gdm.id = gc.municipality_id) org on org.reality_object_id = gro.id
INNER JOIN gkh_dict_work_cur_repair gdwcr on gdwcr.id = gocr.work_kind_id
where inn = "+inn+" and ((plan_date >= to_date('01-01-"+year+"', 'dd-mm-yyyy') and plan_date <= to_date('31-12-"+year+
             "', 'dd-mm-yyyy')) or (fact_date >= to_date('01-01-"+year+"', 'dd-mm-yyyy') and fact_date <= to_date('31-12-"+year+"', 'dd-mm-yyyy'))) order by gro.address, gdwcr.name";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            try
            {
                da.Fill(dt);
                return dt;
            }
            catch (Exception e)
            {
                return null;
            }
            finally
            { conn.Close(); }

        }

        public DataTable SelectLiftInfo()
        {
            string cmdText = @"SELECT gdm.name, gro.address, gro.maximum_floors, gro.number_lifts, t1.value2, t2.value2, t3.value2, t4.value2, t5.value2, t6.value2, t7.value2, 
CASE WHEN t8.value2 is not null THEN to_char(to_date(substr(t8.value2, 1, 10), 'yyyy-mm-dd'), 'dd.mm.yyyy') END as d1
FROM gkh_reality_object gro
inner join gkh_dict_municipality gdm on gdm.id = gro.municipality_id
left join (select ttpv.value as value2, ttp.reality_obj_id as id, ttpv.cell_code
          from tp_teh_passport_value ttpv 
          inner join tp_teh_passport ttp on ttpv.teh_passport_id = ttp.id 
          where ttpv.form_code = 'Form_4_2_1' and ttpv.cell_code LIKE '%:1') t1 on t1.id = gro.id
left join (select ttpv.value as value2, ttp.reality_obj_id as id, ttpv.cell_code 
          from tp_teh_passport_value ttpv 
          inner join tp_teh_passport ttp on ttpv.teh_passport_id = ttp.id 
          where ttpv.form_code = 'Form_4_2_1' and ttpv.cell_code LIKE '%:2') t2 on t2.id = gro.id AND substr(t2.cell_code, 1, 2) = substr(t1.cell_code, 1, 2)  
left join (select ttpv.value as value2, ttp.reality_obj_id as id, ttpv.cell_code
          from tp_teh_passport_value ttpv 
          inner join tp_teh_passport ttp on ttpv.teh_passport_id = ttp.id 
          where ttpv.form_code = 'Form_4_2_1' and ttpv.cell_code LIKE '%:3') t3 on t3.id = gro.id AND substr(t3.cell_code, 1, 2) = substr(t1.cell_code, 1, 2)   
left join (select ttpv.value as value2, ttp.reality_obj_id as id, ttpv.cell_code
          from tp_teh_passport_value ttpv 
          inner join tp_teh_passport ttp on ttpv.teh_passport_id = ttp.id 
          where ttpv.form_code = 'Form_4_2_1' and ttpv.cell_code LIKE '%:9') t4 on t4.id = gro.id AND substr(t4.cell_code, 1, 2) = substr(t1.cell_code, 1, 2)
left join (select ttpv.value as value2, ttp.reality_obj_id as id, ttpv.cell_code
          from tp_teh_passport_value ttpv 
          inner join tp_teh_passport ttp on ttpv.teh_passport_id = ttp.id 
          where ttpv.form_code = 'Form_4_2_1' and ttpv.cell_code LIKE '%:10') t5 on t5.id = gro.id AND substr(t5.cell_code, 1, 2) = substr(t1.cell_code, 1, 2)
left join (select ttpv.value as value2, ttp.reality_obj_id as id, ttpv.cell_code
          from tp_teh_passport_value ttpv 
          inner join tp_teh_passport ttp on ttpv.teh_passport_id = ttp.id 
          where ttpv.form_code = 'Form_4_2_1' and ttpv.cell_code LIKE '%:11') t6 on t6.id = gro.id AND substr(t6.cell_code, 1, 2) = substr(t1.cell_code, 1, 2)
left join (select ttpv.value as value2, ttp.reality_obj_id as id, ttpv.cell_code
          from tp_teh_passport_value ttpv 
          inner join tp_teh_passport ttp on ttpv.teh_passport_id = ttp.id 
          where ttpv.form_code = 'Form_4_2_1' and ttpv.cell_code LIKE '%:12') t7 on t7.id = gro.id AND substr(t7.cell_code, 1, 2) = substr(t1.cell_code, 1, 2)
left join (select ttpv.value as value2, ttp.reality_obj_id as id, ttpv.cell_code
          from tp_teh_passport_value ttpv 
          inner join tp_teh_passport ttp on ttpv.teh_passport_id = ttp.id 
          where ttpv.form_code = 'Form_4_2_1' and ttpv.cell_code LIKE '%:13') t8 on t8.id = gro.id AND substr(t8.cell_code, 1, 2) = substr(t1.cell_code, 1, 2)
where (gro.maximum_floors is null or gro.maximum_floors > 5) and gro.condition_house in (20, 30) and gro.type_house in (30, 40)
order by gdm.name, gro.address, t1.value2";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            try
            {
                da.Fill(dt);
                return dt;
            }
            catch (Exception e)
            {
                return null;
            }
            finally
            { conn.Close(); }

        }

        public DataTable SelectPctInfo()
        {
            string cmdText = @"SELECT org.name as manOrg, CASE WHEN per.percent is not null THEN per.percent ELSE 0 END as pct, org.phone, org.juridical_address, org.fact_address, org2.countHouse
FROM GKH_REALITY_OBJECT gro
left JOIN(SELECT t3.reality_object_id, gc.id, gc.name as name, gc.phone, gc.fact_address, gdm.name || ', ' || gc.juridical_address as juridical_address
                      FROM(select t2.reality_object_id, mo.contragent_id from(SELECT manag_org_id, gmcr.reality_obj_id as reality_object_id, max(gmcr.id) 
                  FROM gkh_morg_contract gmc
                  INNER JOIN gkh_morg_contract_realobj gmcr on gmcr.man_org_contract_id = gmc.id
                  where current_date>start_date AND(current_date < end_date OR end_date is null)
                  group by manag_org_id, gmcr.reality_obj_id) t2
                              left join gkh_managing_organization mo on mo.id = t2.manag_org_id) t3 inner join gkh_contragent gc on gc.id = t3.contragent_id
                              inner join gkh_dict_municipality gdm on gdm.id = gc.municipality_id) org on org.reality_object_id = gro.id
inner join (SELECT count(gro.address) as countHouse, org.id
FROM GKH_REALITY_OBJECT gro
left JOIN(SELECT t3.reality_object_id, gc.id, gc.name as name, gdm.name || ', ' || gc.juridical_address as juridical_address
                      FROM(select t2.reality_object_id, mo.contragent_id from(SELECT manag_org_id, gmcr.reality_obj_id as reality_object_id, max(gmcr.id) 
                  FROM gkh_morg_contract gmc
                  INNER JOIN gkh_morg_contract_realobj gmcr on gmcr.man_org_contract_id = gmc.id
                  where current_date>start_date AND(current_date < end_date OR end_date is null)
                  group by manag_org_id, gmcr.reality_obj_id) t2
                              left join gkh_managing_organization mo on mo.id = t2.manag_org_id) t3 inner join gkh_contragent gc on gc.id = t3.contragent_id
                              inner join gkh_dict_municipality gdm on gdm.id = gc.municipality_id) org on org.reality_object_id = gro.id
where org.name is not null
group by org.id) org2 on org.id = org2.id
left join (SELECT ro_id, percent from gkh_house_prov_passport where rep_month = 1 AND rep_year = 2015) per on per.ro_id = gro.id
order by countHouse desc, org.name, per.percent";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            try
            {
                da.Fill(dt);
                return dt;
            }
            catch (Exception e)
            {
                return null;
            }
            finally
            { conn.Close(); }

        }

        public DataTable SelectDisp()
        {
            string cmdText = @"SELECT id, type_, CASE WHEN gji_number is not null THEN gji_number ELSE inspection_number END as num, document_date, num1, num2
FROM (SELECT gi.id, CASE WHEN gi.type_base = 30 THEN 'Плановая проверка юридических лиц' WHEN gi.type_base = 50 THEN 'Проверки по требованию прокуратуры' 
WHEN gi.type_base = 40 THEN 'Проверка по поручению руководителя' WHEN gi.type_base = 20 THEN 'Проверка обращениям граждан'
WHEN gi.type_base = 110 THEN 'Проверка по плану мероприятий' WHEN gi.type_base = 150 THEN 'Проверка без основания'
WHEN gi.type_base = 10 THEN 'инспекционная проверка' END as type_, 
gi.inspection_number,gd1.document_date,gd1.document_number as num1, gd2.document_number as num2, gac.gji_number
FROM gji_inspection gi
LEFT JOIN gji_document gd1 on gd1.inspection_id = gi.id AND gd1.type_document = 10
LEFT JOIN gji_document gd2 on gd2.inspection_id = gi.id AND gd2.type_document = 20 aND gd1.document_num = gd2.document_num
LEFT join gji_basestat_appcit gba on gba.inspection_id = gi.id
LEFT JOIN gji_appeal_citizens gac on gac.id = gba.gji_appcit_id 
order by gi.id, type_, gd1.document_date) t1
where type_ is not null";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            try
            {
                da.Fill(dt);
                return dt;
            }
            catch (Exception e)
            {
                return null;
            }
            finally
            { conn.Close(); }

        }

        public DataTable SelectInfo(int type)
        {
            string cmdText = @"SELECT id, name, address, email, official_website, CAST(RTRIM(Sys_xmlagg(XMLELEMENT(col, phone||', ')).extract('/ROWSET/COL/text()').getclobval(), ', ') AS VARCHAR2(4000)) as phone, ogrn, address2
from(SELECT gc.id, gc.name, case when bfa.address_name is null then gc.juridical_address 
when bf.postalcode is null then bfa.address_name else bfa.address_name || ',' || bf.postalcode END as address,
case when bfa2.address_name is null then gc.fact_address 
when bf2.postalcode is null then bfa2.address_name else bfa2.address_name || ',' || bf2.postalcode END as address2,
gc.email, gc.official_website, CASE WHEN gc.phone is null THEN gcc.phone ELSE gc.phone END as phone, gc.phone_dispatch_service,
gc.ogrn
FROM gkh_managing_organization gmo
INNER JOIN gkh_contragent gc on gc.id = gmo.contragent_id
LEFT JOIN gkh_contragent_contact gcc on gc.id = gcc.contragent_id
LEFT JOIN b4_fias_address bfa on bfa.id = gc.fias_jur_address_id
LEFT JOIN b4_fias bf on bf.aoguid = bfa.street_guid and bf.postalcode is not null
LEFT JOIN b4_fias_address bfa2 on bfa2.id = gc.fias_fact_address_id
LEFT JOIN b4_fias bf2 on bf2.aoguid = bfa2.street_guid and bf2.postalcode is not null
where gmo.type_management = "+type+@" and gmo.activity_termination = 10) t1
group by id, name, address, address2, email, official_website, ogrn
order by name";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            try
            {
                da.Fill(dt);
                return dt;
            }
            catch (Exception e)
            {
                return null;
            }
            finally
            { conn.Close(); }

        }

        public DataTable SelectInfo()
        {
            string cmdText = @"SELECT id, name, address, email, official_website, CAST(RTRIM(Sys_xmlagg(XMLELEMENT(col, phone||', ')).extract('/ROWSET/COL/text()').getclobval(), ', ') AS VARCHAR2(4000)) as phone, ogrn, address2
from(SELECT gc.id, gc.name, case when bfa.address_name is null then gc.juridical_address 
when bf.postalcode is null then bfa.address_name else bfa.address_name || ',' || bf.postalcode END as address,
case when bfa2.address_name is null then gc.fact_address 
when bf2.postalcode is null then bfa2.address_name else bfa2.address_name || ',' || bf2.postalcode END as address2,
gc.email, gc.official_website, CASE WHEN gc.phone is null THEN gcc.phone ELSE gc.phone END as phone, gc.phone_dispatch_service,
gc.ogrn
FROM gkh_managing_organization gmo
INNER JOIN gkh_contragent gc on gc.id = gmo.contragent_id
LEFT JOIN gkh_contragent_contact gcc on gc.id = gcc.contragent_id
LEFT JOIN b4_fias_address bfa on bfa.id = gc.fias_jur_address_id
LEFT JOIN b4_fias bf on bf.aoguid = bfa.street_guid and bf.postalcode is not null
LEFT JOIN b4_fias_address bfa2 on bfa2.id = gc.fias_fact_address_id
LEFT JOIN b4_fias bf2 on bf2.aoguid = bfa2.street_guid and bf2.postalcode is not null
where gmo.type_management not in (20, 10) and gmo.activity_termination = 10) t1
group by id, name, address, address2, email, official_website, ogrn
order by name";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            try
            {
                da.Fill(dt);
                return dt;
            }
            catch (Exception e)
            {
                return null;
            }
            finally
            { conn.Close(); }

        }

        public string UpdateHouseOwner(string innFrom, string manOrgIdTo)
        {
            string cmdText = "select * from gkh_man_org_real_obj where manag_org_id =(select id from gkh_managing_organization where contragent_id = (SELECT id from gkh_contragent where inn = '"+innFrom+"'))";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            try
            {
                da.Fill(dt);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string cmdText1 = "INSERT INTO gkh_man_org_real_obj (ID, object_version, object_create_date, object_edit_date, manag_org_id, reality_object_id, date_start" +
                       ") VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, '" + manOrgIdTo + "'," + dt.Rows[i][5].ToString() + "," +
                       "TO_TIMESTAMP('01.09.2014', 'DD.MM.YYYY'))";
                    OracleCommand cmd2 = new OracleCommand(cmdText1, conn);
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

        public string InsertPeople3(string gkh_code, string flat, string fio, string total_area)
        {
            string cmdText = "SELECT gro.id from GKH_REALITY_OBJECT gro inner join GKH_DICT_MUNICIPALITY gdm on GDM.ID = GRO.MUNICIPALITY_ID where GDM.NAME LIKE '%г. Самара, Красноглинский р-н' AND replace(lower(GRO.ADDRESS), ' ','') = replace(lower('" + gkh_code + "'), ' ','')";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
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
               
            if (total_area == null || total_area == "" || total_area == " ")
                total_area = "0";
           
            if(!houses.Contains(id.ToString()))
            {
                houses.Add(id.ToString());
                string cmdText2 = "DELETE FROM gkh_obj_apartment_info where reality_object_id = " + id;
                OracleCommand cmd2 = new OracleCommand(cmdText2, conn);
                cmd2.ExecuteNonQuery();
            }

            string cmdText1 = "INSERT INTO gkh_obj_apartment_info (id, object_version, object_create_date, object_edit_date, num_apartment, area_total, privatized," +
                    "reality_object_id, fio_owner) VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, '" + flat + "'," + total_area
                    + "," + priv + "," + id + ",'" + fio + "')";
            OracleCommand cmd1 = new OracleCommand(cmdText1, conn);
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

        public string InsertPeople4(string gkh_code, string flat, string fio, string total_area)
        {
            string cmdText = "SELECT id from GKH_REALITY_OBJECT where gkh_code = " + gkh_code;
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
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

            if (total_area == null || total_area == "" || total_area == " ")
                total_area = "0";

            if (!houses.Contains(id.ToString()))
            {
                houses.Add(id.ToString());
                string cmdText2 = "DELETE FROM gkh_obj_apartment_info where reality_object_id = " + id;
                OracleCommand cmd2 = new OracleCommand(cmdText2, conn);
                cmd2.ExecuteNonQuery();
            }

            string cmdText1 = "INSERT INTO gkh_obj_apartment_info (id, object_version, object_create_date, object_edit_date, num_apartment, area_total, privatized," +
                    "reality_object_id, fio_owner) VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, '" + flat + "'," + total_area
                    + "," + priv + "," + id + ",'" + fio + "')";
            OracleCommand cmd1 = new OracleCommand(cmdText1, conn);
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

        public string InsertPeople4(string gkh_code, string flat, string fio,
            string useful_area, string total_area, string residents_count, string privatized)
        {
            string cmdText = "SELECT id from GKH_REALITY_OBJECT where gkh_code = " + gkh_code;
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
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
            if (privatized == "да")
                priv = 10;
            else
                priv = 20;
            if (useful_area == null || useful_area == "" || useful_area == " ")
                useful_area = "0";
            if (total_area == null || total_area == "" || total_area == " ")
                total_area = "0";
            if (residents_count == null || residents_count == "" || residents_count == " ")
                residents_count = "0";

            if (!houses.Contains(id.ToString()))
            {
                houses.Add(id.ToString());
                string cmdText2 = "DELETE FROM gkh_obj_apartment_info where reality_object_id = " + id;
                OracleCommand cmd2 = new OracleCommand(cmdText2, conn);
                cmd2.ExecuteNonQuery();
            }

            string cmdText1 = "INSERT INTO gkh_obj_apartment_info (id, object_version, object_create_date, object_edit_date, num_apartment, area_total, count_people, privatized," +
                    "reality_object_id, fio_owner, area_living) VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, '" + flat + "'," + total_area
                    + "," + residents_count + "," + priv + "," + id + ",'" + fio + "'," + useful_area + ")";
            OracleCommand cmd1 = new OracleCommand(cmdText1, conn);
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

        public string InsertPeople4(string gkh_code, string flat, string fio,
            string total_area, string privatized, string residents_count)
        {
            string cmdText = "SELECT id from GKH_REALITY_OBJECT where gkh_code = " + gkh_code;
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
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
            if (privatized == "да")
                priv = 10;
            else
                priv = 20;
            if (total_area == null || total_area == "" || total_area == " ")
                total_area = "0";
            if (residents_count == null || residents_count == "" || residents_count == " ")
                residents_count = "0";

            if (!houses.Contains(id.ToString()))
            {
                houses.Add(id.ToString());
                string cmdText2 = "DELETE FROM gkh_obj_apartment_info where reality_object_id = " + id;
                OracleCommand cmd2 = new OracleCommand(cmdText2, conn);
                cmd2.ExecuteNonQuery();
            }

            string cmdText1 = "INSERT INTO gkh_obj_apartment_info (id, object_version, object_create_date, object_edit_date, num_apartment, area_total, count_people, privatized," +
                    "reality_object_id, fio_owner) VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, '" + flat + "'," + total_area
                    + "," + residents_count + "," + priv + "," + id + ",'" + fio + "')";
            OracleCommand cmd1 = new OracleCommand(cmdText1, conn);
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

        public string RemovePeople(string idHouseFrom, string idHouseTo)
        {
            string cmdText = "SELECT num_apartment, area_total, count_people, privatized, reality_object_id, fio_owner, area_living from gkh_obj_apartment_info where reality_object_id = " + idHouseFrom;
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            da.Fill(dt);
            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string cmdText1 = "INSERT INTO gkh_obj_apartment_info (id, object_version, object_create_date, object_edit_date, num_apartment, area_total, count_people, privatized," +
                       "reality_object_id, fio_owner, area_living) VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, '" + dt.Rows[i][0].ToString() + "'," + dt.Rows[i][1].ToString()
                       + "," + dt.Rows[i][2].ToString() + "," + dt.Rows[i][3].ToString() + "," + idHouseTo + ",'" + dt.Rows[i][5].ToString() + "'," + dt.Rows[i][6].ToString() + ")";
                    OracleCommand cmd1 = new OracleCommand(cmdText1, conn);
                    cmd1.ExecuteNonQuery();
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

            //string cmdText2 = "DELETE FROM gkh_obj_apartment_info where reality_object_id = " + id;
            //OracleCommand cmd2 = new OracleCommand(cmdText2, conn);
            //cmd2.ExecuteNonQuery();
            //return "ЗАГРУЖЕНО";
        }

        public DataTable SelectInspection()
        {
            string cmdText = @"SELECT gac.document_number, gdm.name, gro.address, gdz.name, gac.date_from, gac.check_time, gdi.fio, gdi2.fio
FROM gji_appeal_citizens gac
INNER JOIN gji_appcit_ro gar on gar.appcit_id = gac.id
INNER JOIN gkh_reality_object gro on gro.id = gar.reality_object_id
INNER JOIN gkh_dict_municipality gdm on gdm.id = gro.municipality_id
INNER JOIN gkh_dict_zonainsp gdz on gdz.id = gac.zonainsp_id
INNER JOIN gkh_dict_inspector gdi on gdi.id = gac.executant_id
INNER JOIN gji_sam_appcits_tester gsat on gsat.appealcit_id = gac.id
INNER JOIN gkh_dict_inspector gdi2 on gdi2.id = gsat.tester_id
order by gdm.name, gac.document_number";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            conn.Open();
            try
            {
                da.Fill(dt);
                return dt;
            }
            catch (Exception e)
            {
                return null;
            }
            finally
            { conn.Close(); }

        }

        public string InsertRealtyObject(string mo, string address, decimal total_area, int people_count)
        {
            string cmdText = "SELECT id from GKH_DICT_MUNICIPALITY gdm where GDM.NAME LIKE '%" + mo + "%'";
            OracleConnection conn = new OracleConnection(connStr);
            OracleCommand cmd = new OracleCommand(cmdText, conn);
            OracleDataAdapter da = new OracleDataAdapter(cmd);
            DataTable dt = new DataTable();
            int municipality_id;
            conn.Open();
            da.Fill(dt);
            municipality_id = Convert.ToInt32(dt.Rows[0][0]);


            string cmdText1 = "INSERT INTO gkh_reality_object (id, object_version, object_create_date, object_edit_date, municipality_id, type_ownership_id, capital_group_id, condition_house, " + 
                "having_basement, heating_system, type_house, address,area_basement, area_living, area_liv_not_liv_mkd, area_living_owned, area_mkd, floors, maximum_floors, is_insured_object, " + 
                "number_apartments, number_entrances, number_lifts, number_living, physical_wear, residents_evicted, is_build_soc_mortgage, area_owned, area_municipal_owned, necessary_conduct_cr, " 
                + "state_id, method_form_fund, judgment_common_prop, has_priv_flats, project_docs, energy_passport, confirm_work_docs, is_not_involved_cr) " +
                "VALUES(HIBERNATE_SEQUENCE.NEXTVAL, 0, CURRENT_DATE, CURRENT_DATE, " + municipality_id + ", 293823, 293802, 30, 10, 10, 20, '" + address + "', 0,0,0,0," + total_area + 
                ", 1, 1, 0, 1, 1, 0, " + people_count + ", 50, 0, 20, 0, 0, 30, 6057411, 0, 20, 1, 10, 10, 10, 1)";
            OracleCommand cmd1 = new OracleCommand(cmdText1, conn);
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

