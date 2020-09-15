using System.Collections.Generic;
using System.IO;
using Excel.DataTable.Models;

namespace Excel.DataTable
{
    public interface IDataWriter
    {
        Models.DataTable WriteToTable(Stream stream, IEnumerable<FilterSet> filterSets, bool isEditable, RowStyles rowStyle, string sheetName = "");

        void UpdateCells(Stream stream, IEnumerable<CellTemplate> cellTemplates);
    }
}