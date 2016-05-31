using Oracle.DataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1.Database
{
    class BillUploadDataDb
    {
        private string connStr;
        private string connStrTest;
        private string connStrHome;
        public BillUploadDataDb()
        {
            //connStr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.126.128)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=XE)));User Id = HR; Password = test";
            //connStr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.1.30)(PORT=1521))(CONNECT_DATA=(SID=ORCL)));User Id=dbq;Password=dbq;";
            //connStr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=192.168.5.77)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=orcl)));User Id = MJF; Password = ActaNonVerba";
            connStr = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=85.140.61.250)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ezhkh)));User Id = b4_gkh_samara; Password = ACTANONVERBA";
            connStrTest = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=85.140.61.250)(PORT=1521))(CONNECT_DATA=(SERVICE_NAME=ezhkh)));User Id = b4_gkh_samara2; Password = ACTANONVERBA";
            connStrHome = "Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=46.0.13.2)(PORT=1578))(CONNECT_DATA=(SID=orcl)));User Id=dbq;Password=dbq";
        }

        public DataTable SelectHouseCode(string region)
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
                  where tpv.form_code = 'Form_4_1' AND tpv.cell_code = '1:4' AND tpv.value >0) t41 on t41.id = gro.id
LEFT JOIN (SELECT ro.id as id, tpv.value from gkh_reality_object ro
                  inner join tp_teh_passport tp on tp.reality_obj_id = ro.id
                  inner join tp_teh_passport_value tpv on tpv.teh_passport_id = tp.id
                  inner join gkh_dict_municipality gdm on gdm.id = ro.municipality_id
                  where tpv.form_code = 'Form_3_7' AND tpv.cell_code = '1:3') t5 on t5.id = gro.id) tab) t1 on t1.id = gro.id
LEFT JOIN (select gro.id, MAX(bf.kladrcode) as kladrcode from gkh_reality_object gro inner join b4_fias_address bfa on bfa.id = gro.fias_address_id 
inner join b4_fias bf on bf.aoguid = bfa.street_guid group by gro.id) t2 on t2.id = gro.id
                            where type_house in(0, 10, 20, 30, 40) AND condition_house !=40 AND (gkh_code =" + region + ") AND LENGTH(bfa.house) < 7 order by gro.gkh_code";
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

        public DataTable SelectLN4Code(string region)
        {
            string cmdText = @"SELECT '4' as code, gro.gkh_code as house_code,0 as owner_code,  '1' as typeLS, goai.fio_owner, goai.num_apartment, goai.count_people, '0' as COunt17, '0' as count18, '1' as count19, goai.area_total,
 goai.area_living, '0' as count24, CASE WHEN goai.privatized = 20 THEN 0 ELSE 1 END as qwerty
 FROM gkh_obj_apartment_info goai
 inner join gkh_reality_object gro on gro.id = goai.reality_object_id
 where reality_object_id in (
 SELECT gro.id from gkh_reality_object gro
 LEFT JOIN (select gro.id, MAX(bf.kladrcode) as kladrcode from gkh_reality_object gro inner join b4_fias_address bfa on bfa.id = gro.fias_address_id 
inner join b4_fias bf on bf.aoguid = bfa.street_guid where LENGTH(bfa.house) < 7 group by gro.id) t2 on t2.id = gro.id 
where type_house not in(0, 20) AND condition_house !=40 AND gkh_code = " + region + ") order by house_code desc";
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
    }
}
