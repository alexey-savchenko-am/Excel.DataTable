using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using DataHandler.Excel.Models;

namespace DataHandler.Excel
{
    public interface IDataObtainer
    {
        DataTable ObtainTable(string filePath, IEnumerable<FilterValue> filterValues, bool isEditable,
            string sheetName = "");

        Task<DataTable> ObtainTableAsync(Stream stream, IEnumerable<FilterValue> filterValues, bool isEditable,
            string sheetName = "");

        DataTable ObtainTable(Stream stream, IEnumerable<FilterValue> filterValues, bool isEditable,
            bool disposeStreamAfterReading = true, string sheetName = "");
    }
}