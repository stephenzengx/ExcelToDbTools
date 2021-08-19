using System;
using Dapper;
using System.Collections.Generic;

namespace ExcelTools
{
    /// <summary>
    /// 读取Excel表头结果
    /// </summary>
    public class ScanExcelHeadDesc
    {
        /// <summary>
        /// 前缀
        /// </summary>
        public string Prefix { get; set; } = string.Empty;

        /// <summary>
        /// Excel表头名
        /// </summary>
        public string HeaderName { get; set; }

        /// <summary>
        /// 对应数据库字段名
        /// </summary>
        public string FieldName { get; set; }

        /// <summary>
        /// Decimal范围
        /// </summary>
        public Tuple<decimal, decimal> RangeDecimal { get; set; } = new Tuple<decimal, decimal>(0,decimal.MaxValue);

        /// <summary>
        /// int范围
        /// </summary>
        public Tuple<int, int> RangeInt { get; set; } = new Tuple<int, int>(0, int.MaxValue);

        /// <summary>
        /// 关联表名
        /// </summary>
        public string RelatedTbName { get; set; }

        /// <summary>
        /// 所关联的字段
        /// </summary>
        public string KeyFieldName { get; set; }

        //public object KeyFieldValue { get; set; }

        /// <summary>
        /// 通过关联字段要取到的另一个字段值
        /// </summary>
        public string ValueFieldName { get; set; }
    }

    /// <summary>
    /// 读取Excel表格数据结果 
    /// </summary>
    public class ReadExcelDataRet
    {
        public string DbName { get; set; }
        public string TbName { get; set; }

        public string ExecSql { get; set; } // sql

        public List<string> FieldNameList { get; set; } 

        public List<DynamicParameters> ExecParams { get; set; } //sql 参数

        public List<int> RowIndexs { get; set; }

        public int AllRowCount { get; set; } //验证总条数

        public int LegalRowCount { get; set; } //验证通过条数

        public int IllegalRowCount { get; set; } //验证不通过条数
    }

    /// <summary>
    /// 数据库-表属性
    /// </summary>
    public class TbDesc
    {
        public string FullTbName { get; set; }
        public string FieldName { get; set; }
        public string FieldType { get; set; }
        public bool IsNeeded { get; set; }
        public Type type { get; set; }//字段类型对应C# 基础数据类型

        public int MaxLength { get; set; } //如果为varchar的最大长度
    }

    public class RelatedDicClass
    {
        public string RKey { get; set; }
        public string RValue { get; set; }
    }
    public class HospitalInfo
    {
        public string Id { get; set; }
        public string Mc { get; set; }
    }

    public class RltNoMatchInfo
    {
        public int Sort { get; set; }
        public string Info { get; set; }

        public RltNoMatchInfo(int sort, string info)
        {
            Sort = sort;
            Info = info;
        }
    }

    /// <summary>
    /// 标识符枚举
    /// </summary>
    public enum EnumIdentifier
    {
        Dot,
        LeftZkh,//左中括号
        RightZkh,//右中括号
        Comma,//逗号
        SpitChar,//表头字段分隔符
        Unique,//唯一
        PYWB, //拼音五笔
        Encry, //加密
        Range, //范围
        Related, //关联
        Password //密码
    }
}
