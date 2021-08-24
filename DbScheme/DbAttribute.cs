using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelTools
{
    [AttributeUsage(AttributeTargets.Class)]
    public class DbAttribute : Attribute
    {
        public string DbName { get; set; }

        public DbAttribute(string dbName = "ForExcelDb")
        {
            DbName = dbName;
        }
    }
}
