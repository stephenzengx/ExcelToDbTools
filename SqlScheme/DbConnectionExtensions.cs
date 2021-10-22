using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Dapper;
//using ExcelTools.DbScheme;
using ExcelTools.SqlScheme;
//using ForExcelImport;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using MySql.Data.MySqlClient;

namespace ExcelTools
{
    public static class DbConnectionExtensions
    {
        /// <summary>
        /// 批量修改 手机号/身份证号
        /// </summary>
        /// <param name="con"></param>
        /// <param name="list"></param>
        //public static void Change_SJHM_ZJHM(this IDbConnection con,List<tb_relative> list)
        //{
        //    var i = 1;
        //    foreach (var item in list)
        //    {
        //        DynamicParameters sParameters = new DynamicParameters();

        //        sParameters.Add("xm", item.xm);
        //        sParameters.Add("sjhm", item.sjhm.ToPhoneCardNoEncryption());
        //        sParameters.Add("zjhm", item.zjhm.ToPhoneCardNoEncryption());

        //        var count = con.Execute($"update t_base_user set sjhm = @sjhm, zjhm=@zjhm where xm=@xm", sParameters);
        //        i++;
        //        if (count > 0)
        //            Console.WriteLine($"{i}-{item.xm} 修改成功");
        //    }
        //}

        /// <summary>
        /// 批量修改密码
        /// </summary>
        public static void ChangeMM()
        {
            var dbname = "ihbase";
            using (IDbConnection con = new MySqlConnection(Utils.Config.GetConnectionString(dbname)))
            {
                con.Open();

                DynamicParameters Parameters = new DynamicParameters();
                Parameters.Add("jgid", "c28c7004-9f80-4460-8510-b626710e3b24");

                var dbsql = @"select id from t_base_user where yljguuid=@jgid"; //__efmigrationshistory
                var ids = con.Query<string>(dbsql, Parameters);

                foreach (var id in ids)
                {
                    var password = PwdHelper.ToPassWordGetSalt("Hlwyy@135", out string salt);

                    DynamicParameters p = new DynamicParameters();
                    p.Add("mm", password);
                    p.Add("mmy", salt);

                    var uSql = $"update t_base_user set mm =@mm,mmy=@mmy ";
                    con.Execute(uSql, p);
                }
            }
        }

        /// <summary>
        /// 测试药品 单位转换
        /// </summary>
        //public static void TestDrugCalulate()
        //{
        //    using (DbContext db = DbSchemeHelper.GetTestDbContext(typeof(TestDbContext)))
        //    {
        //        var drug = db.Set<Drug>().FirstOrDefault(m => m.id == "f9c85319-7f16-4874-b42a-8ca1583b588b");

        //        //总量取整
        //        var customeUnit = new CustomeUnit($"{drug.ZXDW}/{drug.JLDW}/{drug.BZDW}", $"1/{drug.JLSL}/{drug.BZSL}");

        //        var num = (int)Math.Ceiling(customeUnit.UnitConversion(2, "把", drug.YLDW));
        //    }
        //}

        /// <summary>
        /// 批量更新 cjry,xgry
        /// </summary>
        public static void UpdeteBatchRYID()
        {
            var list = new List<string> { "ihbase" }; //"ihbase_lgfy" ,"ihdb_lgfy" ,"cloudih", "cloudih" 

            foreach (var dbname in list)
            {
                UpdateBatch(dbname);
            }
        }

        public static void UpdateBatch(string dbname)
        {
            using (IDbConnection con = new MySqlConnection(Utils.Config.GetConnectionString(dbname)))
            {
                con.Open();

                DynamicParameters Parameters = new DynamicParameters();
                Parameters.Add("dbname", dbname);

                var dbsql = @"select LOWER(table_name) from information_schema.tables where table_schema=@dbname"; //__efmigrationshistory
                var tbNameList = con.Query<string>(dbsql, Parameters);

                foreach (var tbName in tbNameList)
                {
                    try
                    {
                        DynamicParameters p = new DynamicParameters();
                        //p.Add("cjry", "00000000-0000-0000-0000-000000000000");
                        //p.Add("xgry", "00000000-0000-0000-0000-000000000000");
                        p.Add("xgsj", DateTime.Now);



                        var uSql = $"update {tbName} set xgsj=@xgsj"; // cjry = @cjry,xgry=@xgry
                        con.Execute(uSql, p);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                    }
                }
            }
        }

        //public static void Change_YZ_PYMWBM(this IDbConnection con, List<tb_relative> list)
        //{
        //    var i = 1;
        //    foreach (var item in list)
        //    {
        //        DynamicParameters sParameters = new DynamicParameters();

        //        sParameters.Add("pym", item.mc.GetFirstPY());
        //        sParameters.Add("wbm", item.mc.GetFirstWB());
        //        sParameters.Add("id", item.id);

        //        var count = con.Execute($"update t_base_medicalorder set pym = @pym, wbm=@wbm where id=@id", sParameters);
        //        i++;
        //        if (count > 0)
        //            Console.WriteLine($"{i}-{item.mc} 修改成功");
        //    }
        //}

        public static void BulkInsert(this IDbConnection con, string execSql, List<DynamicParameters> pParameters)
        {
            con.Execute(execSql, pParameters); //直接传送list对象
        }

        public static List<string> CommonFieldNames = new List<string> { "id", "isdeleted", "cjsj", "cjry","cjrymc","xgsj","xgry","xgrymc"};

        public static List<string> PYWBS = new List<string>{"pym","wbm"};

        /// <summary>
        /// 插入sql主语句拼接 (不包括参数)
        /// </summary>
        /// <param name="tbname"></param>
        /// <param name="fieldNames"></param>
        /// <param name="isPyWB"></param>
        /// <param name="JGIDFieldName"></param>
        /// <param name="existFieldNames"></param>
        /// <returns></returns>
        public static string JoinInsertHeadSql(string dbname, string tbname, List<string> fieldNames, bool isPyWB, bool isExistJybs, string JGIDFieldName,  out List<string> existFieldNames)
        {
            StringBuilder sb = new StringBuilder();

            fieldNames.Add(JGIDFieldName);

            fieldNames.AddRange(CommonFieldNames);
            
            if (isPyWB)
            {
                fieldNames.AddRange(PYWBS);
            }

            if (isExistJybs)
            {
                fieldNames.Add("jybs");
            }

            sb.Append($" INSERT INTO {dbname}.{tbname} ({string.Join(",", fieldNames)}) ");

            for (int i = 0; i < fieldNames.Count; i++)
            {
                fieldNames[i] = "@" + fieldNames[i];
            }

            sb.Append($"values ({string.Join(",", fieldNames)})");
            
            existFieldNames = fieldNames;

            return sb.ToString();
        }

        /// <summary>
        /// 加入公共参数 值  / parm.Add("JYBS", 0); [改字段不是公共字段 有的地方是int，有的是bool 看下会不会报错]
        /// </summary>
        /// <param name="JGIDFiledName"></param>
        /// <param name="JGID"></param>
        /// <returns></returns>
        public static DynamicParameters GetBaseDynamicParameters(string JGIDFiledName, string JGID,bool isExistJybs)
        {
            DynamicParameters parm = new DynamicParameters();
            parm.Add("id", Guid.NewGuid().ToString());
            if (string.IsNullOrEmpty(JGIDFiledName))
            {
                throw  new Exception("获取JGIDFiledName出错，请检查配置");
            }

            parm.Add(JGIDFiledName, JGID);
            parm.Add("isdeleted", 0);

            parm.Add("cjry", "00000000-0000-0000-0000-000000000000");
            parm.Add("cjrymc", string.Empty);
            parm.Add("cjsj", DateTime.Now);

            parm.Add("xgry", "00000000-0000-0000-0000-000000000000");
            parm.Add("xgrymc", string.Empty);
            parm.Add("xgsj", DateTime.Now);

            if (isExistJybs)
                parm.Add("jybs", false);

            return parm;
        }

        public static List<TbDesc> GetTbDesc(this IDbConnection con, string dbname, string tbname)
        {
            DynamicParameters Parameters = new DynamicParameters();
            Parameters.Add("dbname", dbname);
            Parameters.Add("tbname", tbname);

            string tbsql = @"select LOWER(column_name) as FieldName, LOWER(column_type) as FieldType,  (  case when is_nullable != 'NO' then 0 when is_nullable = 'NO' then 1 end  ) as 'IsNeeded' "
                           + "from information_schema.columns where table_schema=@dbname and table_name = @tbname;";
            var descList = con.Query<TbDesc>(tbsql, Parameters).ToList();
            if (!descList.Any())
                return descList;

            for (var i=0;i<descList.Count; i++)
            {
                descList[i].FullTbName = $"{dbname}.{tbname}";
                var dbType = descList[i].FieldType;
                if (dbType.Equals("longtext"))
                {
                    descList[i].type = typeof(string);
                }
                else if (dbType.StartsWith("varchar"))
                {
                    descList[i].type = typeof(string);
                    var length = dbType.LastIndexOf(")") - dbType.LastIndexOf("(")-1;
                    descList[i].MaxLength = dbType.Substring(8, length).ToInt32();
                }
                else if (dbType.StartsWith("decimal"))
                {
                    descList[i].type = typeof(decimal);
                }
                else if (dbType.StartsWith("int"))
                {
                    descList[i].type = typeof(int);
                }
                else if (dbType.StartsWith("bit") || dbType.StartsWith("tinyint"))
                {
                    descList[i].type = typeof(bool);
                }
                else if (dbType.StartsWith("datetime") || dbType.StartsWith("date"))
                {
                    descList[i].type = typeof(DateTime);
                }
            }

            return descList;
        }

        public static Dictionary<string, List<TbDesc>> GetTableNamesAndFieldDic(this IDbConnection con, string dbname,ref Dictionary<string, List<TbDesc>> dic)
        {
            dbname = dbname.ToLower();
            DynamicParameters Parameters = new DynamicParameters();
            Parameters.Add("dbname", dbname);

            var dbsql = @"select LOWER(table_name) from information_schema.tables where table_schema=@dbname"; //__efmigrationshistory
            var tbNameList = con.Query<string>(dbsql, Parameters);
            if (tbNameList.Count() <= 0)
            {
                Utils.LogInfo("该数据库不存在数据库表,请检查");
                return null;
            }

            foreach (var tbName in tbNameList)
            {
                if(!dic.TryGetValue($"{dbname}.{tbName}",out var _v2))
                    dic.Add($"{dbname}.{tbName}", con.GetTbDesc(dbname, tbName));
            }

            return dic;
        }

        public static List<string> GetUniqFieldValues(this IDbConnection con, string tbName, string fieldName,string JGIDFieldName, string JGID)
        {
            string tbsql = $"select distinct {fieldName} from {tbName}  where {JGIDFieldName} = @JGID;";
            DynamicParameters parameters = new DynamicParameters();
            parameters.Add("JGID", JGID);

            return con.Query<string>(tbsql, parameters).ToList();
        }

        /// <summary>
        /// 获取 key-values 字典
        /// </summary>
        /// <param name="con"></param>
        /// <param name="fullTbName"></param>
        /// <param name="keyFieldName"></param>
        /// <param name="valueFieldName"></param>
        /// <param name="JGIDFieldName"></param>
        /// <param name="JGID"></param>
        /// <returns></returns>
        public static Dictionary<string,string> GetDicValues(this IDbConnection con, string fullTbName, List<string> keyFieldNameList, string valueFieldName, string JGIDFieldName, string JGID)
        {
            var dic = new Dictionary<string, string>();
            if (keyFieldNameList.Count <= 0)
            {
                throw  new ArgumentException("keyFieldNameList error!");
            }

            string tbsql = $"select CONCAT({string.Join(",',',", keyFieldNameList)}) as RKey, {valueFieldName} as RValue from {fullTbName} ";
            if (!fullTbName.Contains("t_base_hospital"))
            {
                tbsql += $"where {JGIDFieldName} = @JGID;";
            }

            DynamicParameters parameters = new DynamicParameters();
            parameters.Add("JGID", JGID);
            var sqlRet = con.Query<RelatedDicClass>(tbsql, parameters).ToList();

            if (sqlRet.Count <= 0)
                return dic;

            sqlRet.ForEach(p =>
            {
                if (!dic.TryGetValue(p.RKey, out var _v2))
                    dic.Add(p.RKey, p.RValue);
            });

            return dic;
        }
    }
}
