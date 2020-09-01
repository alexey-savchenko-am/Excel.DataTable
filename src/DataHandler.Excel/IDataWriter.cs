using System.Collections.Generic;
using System.IO;
using DataHandler.Excel.Models;

namespace DataHandler.Excel
{
    public interface IDataWriter
    {
        DataTable WriteToTable(Stream stream, IEnumerable<FilterSet> filterSets, bool isEditable, RowStyles rowStyle, string sheetName = "");

        void UpdateCells(Stream stream, IEnumerable<CellTemplate> cellTemplates);
    }
}