using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Excel
{
    internal interface IExcelReader
    {
        IEnumerable<IDictionary<string, object>> Query(bool UseHeaderRow, string sheetName,string startCell, IConfiguration configuration);
        IEnumerable<T> Query<T>(string sheetName, string startCell, IConfiguration configuration) where T : class, new();

        /// <summary>
        /// 新增
        /// </summary>
        /// <param name="type"></param>
        /// <param name="sheetName"></param>
        /// <param name="startCell"></param>
        /// <param name="configuration"></param>
        /// <returns></returns>
        IEnumerable<object> Query(Type type, string sheetName, string startCell, IConfiguration configuration);
    }

    internal interface IExcelReaderAsync : IExcelReader
    {
        Task<IEnumerable<IDictionary<string, object>>> QueryAsync(bool UseHeaderRow, string sheetName, string startCell, IConfiguration configuration);
        Task<IEnumerable<T>> QueryAsync<T>(string sheetName, string startCell, IConfiguration configuration) where T : class, new();
    }
}
