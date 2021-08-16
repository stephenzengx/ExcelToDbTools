
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel.OpenXml;
using Excel.Utils;
using Excel.Zip;
using ExcelTools;

namespace Excel
{
    public static class MyMiniExcel
    {
        public static Task SaveAsAsync(
          string path,
          object value,
          bool printHeader = true,
          string sheetName = "Sheet1",
          ExcelType excelType = ExcelType.UNKNOWN,
          IConfiguration configuration = null)
        {
            return Task.Run((Action)(() => SaveAs(path, value, printHeader, sheetName, excelType, configuration)));
        }

        public static Task SaveAsAsync(
          this Stream stream,
          object value,
          bool printHeader = true,
          string sheetName = "Sheet1",
          ExcelType excelType = ExcelType.XLSX,
          IConfiguration configuration = null)
        {
            return GetWriterProvider(stream, sheetName, excelType).SaveAsAsync(value, sheetName, printHeader, configuration);
        }

        public static Task<IEnumerable<object>> QueryAsync(
          string path,
          bool useHeaderRow = false,
          string sheetName = null,
          ExcelType excelType = ExcelType.UNKNOWN,
          string startCell = "A1",
          IConfiguration configuration = null)
        {
            return Task.Run<IEnumerable<object>>((Func<IEnumerable<object>>)(() => Query(path, useHeaderRow, sheetName, excelType, startCell, configuration)));
        }

        public static Task<IEnumerable<T>> QueryAsync<T>(
          this Stream stream,
          string sheetName = null,
          ExcelType excelType = ExcelType.UNKNOWN,
          string startCell = "A1",
          IConfiguration configuration = null)
          where T : class, new()
        {
            return ExcelReaderFactory.GetProvider(stream, Excel.Utils.ExcelTypeHelper.GetExcelType(stream, excelType)).QueryAsync<T>(sheetName, startCell, configuration);
        }

        public static Task<IEnumerable<T>> QueryAsync<T>(
          string path,
          string sheetName = null,
          ExcelType excelType = ExcelType.UNKNOWN,
          string startCell = "A1",
          IConfiguration configuration = null)
          where T : class, new()
        {
            return Task.Run<IEnumerable<T>>((Func<IEnumerable<T>>)(() => Query<T>(path, sheetName, excelType, startCell, configuration)));
        }

        public static Task<IEnumerable<IDictionary<string, object>>> QueryAsync(
          this Stream stream,
          bool useHeaderRow = false,
          string sheetName = null,
          ExcelType excelType = ExcelType.UNKNOWN,
          string startCell = "A1",
          IConfiguration configuration = null)
        {
            return GetReaderProvider(stream, excelType).QueryAsync(useHeaderRow, sheetName, startCell, configuration);
        }

        public static Task SaveAsByTemplateAsync(
          this Stream stream,
          string templatePath,
          object value)
        {
            return ExcelTemplateFactory.GetProvider(stream, ExcelType.XLSX).SaveAsByTemplateAsync(templatePath, value);
        }

        public static Task SaveAsByTemplateAsync(
          this Stream stream,
          byte[] templateBytes,
          object value)
        {
            return ExcelTemplateFactory.GetProvider(stream, ExcelType.XLSX).SaveAsByTemplateAsync(templateBytes, value);
        }

        public static Task SaveAsByTemplateAsync(string path, string templatePath, object value)
        {
            return Task.Run((Action)(() => SaveAsByTemplate(path, templatePath, value)));
        }

        public static Task SaveAsByTemplateAsync(string path, byte[] templateBytes, object value)
        {
            return Task.Run((Action)(() => SaveAsByTemplate(path, templateBytes, value)));
        }

        public static Task<DataTable> QueryAsDataTableAsync(
          string path,
          bool useHeaderRow = true,
          string sheetName = null,
          ExcelType excelType = ExcelType.UNKNOWN,
          string startCell = "A1",
          IConfiguration configuration = null)
        {
            return Task.Run<DataTable>((Func<DataTable>)(() => QueryAsDataTable(path, useHeaderRow, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), startCell, configuration)));
        }

        public static Task<DataTable> QueryAsDataTableAsync(
          this Stream stream,
          bool useHeaderRow = true,
          string sheetName = null,
          ExcelType excelType = ExcelType.UNKNOWN,
          string startCell = "A1",
          IConfiguration configuration = null)
        {
            return Task.Run<DataTable>((Func<DataTable>)(() => ExcelOpenXmlSheetReader.QueryAsDataTableImpl(stream, useHeaderRow, ref sheetName, excelType, startCell, configuration)));
        }

        private static IExcelWriterAsync GetWriterProvider(
          Stream stream,
          string sheetName,
          ExcelType excelType)
        {
            if (string.IsNullOrEmpty(sheetName))
                throw new InvalidDataException("Sheet name can not be empty or null");
            if (excelType == ExcelType.UNKNOWN)
                throw new InvalidDataException("Please specify excelType");
            return ExcelWriterFactory.GetProvider(stream, excelType);
        }

        private static IExcelReaderAsync GetReaderProvider(
          Stream stream,
          ExcelType excelType)
        {
            return ExcelReaderFactory.GetProvider(stream, ExcelTypeHelper.GetExcelType(stream, excelType));
        }

        public static void SaveAs(
          string path,
          object value,
          bool printHeader = true,
          string sheetName = "Sheet1",
          ExcelType excelType = ExcelType.UNKNOWN,
          IConfiguration configuration = null)
        {
            if (Path.GetExtension(path).ToLowerInvariant() == ".xlsm")
                throw new NotSupportedException("MiniExcel SaveAs not support xlsm");
            using (FileStream stream = new FileStream(path, FileMode.CreateNew))
                stream.SaveAs(value, printHeader, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), configuration);
        }

        public static void SaveAs(
          this Stream stream,
          object value,
          bool printHeader = true,
          string sheetName = "Sheet1",
          ExcelType excelType = ExcelType.XLSX,
          IConfiguration configuration = null)
        {
            GetWriterProvider(stream, sheetName, excelType).SaveAs(value, sheetName, printHeader, configuration);
        }

        public static IEnumerable<T> Query<T>(
          string path,
          string sheetName = null,
          ExcelType excelType = ExcelType.UNKNOWN,
          string startCell = "A1",
          IConfiguration configuration = null)
          where T : class, new()
        {
            using (FileStream stream = Helpers.OpenSharedRead(path))
            {
                foreach (T obj in stream.Query<T>(sheetName, ExcelTypeHelper.GetExcelType(path, excelType), startCell, configuration))
                    yield return obj;
            }
        }

        public static IEnumerable<T> Query<T>(
          this Stream stream,
          string sheetName = null,
          ExcelType excelType = ExcelType.UNKNOWN,
          string startCell = "A1",
          IConfiguration configuration = null)
            where T : class, new()
        {
            return ExcelReaderFactory.GetProvider(stream, ExcelTypeHelper.GetExcelType(stream, excelType)).Query<T>(sheetName, startCell, configuration);
        }

        #region 新增Type方法
        public static IEnumerable<object> Query(
            string path,
            Type type,
            string sheetName = null,
            ExcelType excelType = ExcelType.UNKNOWN,
            string startCell = "A1",
            IConfiguration configuration = null)
        {
            using (FileStream stream = Helpers.OpenSharedRead(path))
            {
                return stream.Query(type, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), startCell, configuration);
                //foreach (object obj in )
                //    yield return obj;
            }
        }

        public static IEnumerable<object> Query(
            this Stream stream,
            Type type,
            string sheetName = null,
            ExcelType excelType = ExcelType.UNKNOWN,
            string startCell = "A1",
            IConfiguration configuration = null)
        {
            return ExcelReaderFactory.GetProvider(stream, ExcelTypeHelper.GetExcelType(stream, excelType)).Query(type, sheetName, startCell, configuration); //ExcelOpenXmlSheetReader
        }

        #endregion

        public static IEnumerable<object> Query(
          string path,
          bool useHeaderRow = false,
          string sheetName = null,
          ExcelType excelType = ExcelType.UNKNOWN,
          string startCell = "A1",
          IConfiguration configuration = null)
        {
            using (FileStream stream = Helpers.OpenSharedRead(path))
            {
                foreach (object obj in stream.Query(useHeaderRow, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), startCell, configuration))
                    yield return obj;
            }
        }

        public static IEnumerable<object> Query(
          this Stream stream,
          bool useHeaderRow = false,
          string sheetName = null,
          ExcelType excelType = ExcelType.UNKNOWN,
          string startCell = "A1",
          IConfiguration configuration = null)
        {
            return (IEnumerable<object>)ExcelReaderFactory.GetProvider(stream, ExcelTypeHelper.GetExcelType(stream, excelType)).Query(useHeaderRow, sheetName, startCell, configuration);
        }

        public static void SaveAsByTemplate(string path, string templatePath, object value)
        {
            using (FileStream stream = File.Create(path))
                stream.SaveAsByTemplate(templatePath, value);
        }

        public static void SaveAsByTemplate(string path, byte[] templateBytes, object value)
        {
            using (FileStream stream = File.Create(path))
                stream.SaveAsByTemplate(templateBytes, value);
        }

        public static void SaveAsByTemplate(this Stream stream, string templatePath, object value)
        {
            ExcelTemplateFactory.GetProvider(stream, ExcelType.XLSX).SaveAsByTemplate(templatePath, value);
        }

        public static void SaveAsByTemplate(this Stream stream, byte[] templateBytes, object value)
        {
            ExcelTemplateFactory.GetProvider(stream, ExcelType.XLSX).SaveAsByTemplate(templateBytes, value);
        }

        public static DataTable QueryAsDataTable(
          string path,
          bool useHeaderRow = true,
          string sheetName = null,
          ExcelType excelType = ExcelType.UNKNOWN,
          string startCell = "A1",
          IConfiguration configuration = null)
        {
            using (FileStream stream = Helpers.OpenSharedRead(path))
                return stream.QueryAsDataTable(useHeaderRow, sheetName, ExcelTypeHelper.GetExcelType(path, excelType), startCell, configuration);
        }

        public static DataTable QueryAsDataTable(
          this Stream stream,
          bool useHeaderRow = true,
          string sheetName = null,
          ExcelType excelType = ExcelType.UNKNOWN,
          string startCell = "A1",
          IConfiguration configuration = null)
        {
            return ExcelOpenXmlSheetReader.QueryAsDataTableImpl(stream, useHeaderRow, ref sheetName, excelType, startCell, configuration);
        }

        public static List<string> GetSheetNames(string path)
        {
            using (FileStream stream = Helpers.OpenSharedRead(path))
                return stream.GetSheetNames();
        }

        public static List<string> GetSheetNames(this Stream stream)
        {
            return ExcelOpenXmlSheetReader.GetWorkbookRels(new ExcelOpenXmlZip(stream).entries).Select(s => s.Name).ToList();
        }

        public static ICollection<string> GetColumns(
          string path,
          bool useHeaderRow = false,
          string sheetName = null,
          ExcelType excelType = ExcelType.UNKNOWN,
          string startCell = "A1",
          IConfiguration configuration = null)
        {
            using (FileStream stream = Helpers.OpenSharedRead(path))
                return stream.GetColumns(useHeaderRow, sheetName, excelType, startCell, configuration);
        }

        public static ICollection<string> GetColumns(
          this Stream stream,
          bool useHeaderRow = false,
          string sheetName = null,
          ExcelType excelType = ExcelType.UNKNOWN,
          string startCell = "A1",
          IConfiguration configuration = null)
        {
            return (stream.Query(useHeaderRow, sheetName, excelType, startCell, configuration).FirstOrDefault<object>() as IDictionary<string, object>)?.Keys;
        }
    }
}
