using System.Threading.Tasks;

namespace Excel
{
    internal interface IExcelWriter 
    {
        void SaveAs(object value,string sheetName, bool printHeader, IConfiguration configuration);
    }

    internal interface IExcelWriterAsync : IExcelWriter
    {
        Task SaveAsAsync(object value, string sheetName, bool printHeader, IConfiguration configuration);
    }
}
