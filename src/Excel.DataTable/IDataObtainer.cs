using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Excel.DataTable.Models;

namespace Excel.DataTable
{
    public interface IDataObtainer
    {
        Models.DataTable ObtainTable(string filePath, IEnumerable<FilterValue> filterValues, bool isEditable,
            string sheetName = "");

        Task<Models.DataTable> ObtainTableAsync(Stream stream, IEnumerable<FilterValue> filterValues, bool isEditable,
            string sheetName = "");

        Models.DataTable ObtainTable(Stream stream, IEnumerable<FilterValue> filterValues, bool isEditable,
            bool disposeStreamAfterReading = true, string sheetName = "");
    }
}