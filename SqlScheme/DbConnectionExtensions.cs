using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Dapper;

namespace ExcelTools
{
    public static class DbConnectionExtensions
    {
        public static void TestBulkInsert(this IDbConnection con)
        {
            List<DynamicParameters> pParameters = new List<DynamicParameters>();
            string sql = @" INSERT INTO tb_relative (ID,MC,BM)
                            VALUES(@ID,@MC,@BM); ";

            for (int i = 0; i < 2; i++)
            {
                DynamicParameters sParameters = new DynamicParameters();

                sParameters.Add("ID", "id-" + i);
                //sParameters.Add("ID", "id-" + i+1);

                sParameters.Add("MC", "MC-" + i);
                sParameters.Add("BM", "BM-" + i);
                pParameters.Add(sParameters);
            }

            //var ret = pParameters.FirstOrDefault(o => o.Get<string>("MC").Equals("MC-1"));
            con.Execute(sql, pParameters); //直接传送list对象
        }
        public static void BulkInsert(this IDbConnection con, string execSql, List<DynamicParameters> pParameters)
        {
            con.Execute(execSql, pParameters); //直接传送list对象
        }

        public static List<string> CommonFieldNames = new List<string> { "id", "isdeleted", "cjsj", "cjry","cjrymc"};

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
        public static string JoinInsertHeadSql(string tbname, List<string> fieldNames, bool isPyWB, bool isExistJybs, string JGIDFieldName,  out List<string> existFieldNames)
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

            sb.Append($" INSERT INTO {tbname} ({string.Join(",", fieldNames)}) ");

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
            parm.Add(JGIDFiledName, JGID);
            parm.Add("isdeleted", 0);
            parm.Add("cjrymc", "cjry-mc");
            parm.Add("cjry", "cjry-id");
            parm.Add("cjsj", DateTime.Now);

            if(isExistJybs)
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
                else if (dbType.StartsWith("datetime"))
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

        public static Dictionary<string,string> GetDicValues(this IDbConnection con, string fullTbName, string keyFieldName, string valueFieldName, string JGIDFieldName, string JGID)
        {
            if (fullTbName.Equals("ihdb_lgfy.t_base_presthreetype"))
            {
                Console.WriteLine();
            }

            var dic = new Dictionary<string, string>();
            string tbsql = $"select {keyFieldName} as RKey, {valueFieldName} as RValue from {fullTbName} where {JGIDFieldName} = @JGID;";

            DynamicParameters parameters = new DynamicParameters();
            parameters.Add("JGID", JGID);
            var sqlRet = con.Query<RelatedDicClass>(tbsql, parameters).ToList();

            if (sqlRet.Count <= 0)
                return dic;

            sqlRet.ForEach(p =>
            {
                dic.Add(p.RKey, p.RValue);
            });

            return dic;
        }
    }
}
