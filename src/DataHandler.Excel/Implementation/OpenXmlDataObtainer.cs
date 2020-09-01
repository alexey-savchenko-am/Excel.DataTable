using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DataHandler.Excel.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Tools.Excel.Implementation;

namespace DataHandler.Excel.Implementation
{
    public class OpenXmlDataObtainer
        : OpenXmlDataProcessor, IDataObtainer
    {
        
        public DataTable ObtainTable(string filePath,  IEnumerable<FilterValue> filterValues, bool isEditable, string sheetName = "")
        {

            var dataTable = new DataTable();
            
            using (var document = SpreadsheetDocument.Open(filePath, isEditable))
            {
                dataTable.DataRows = ExtractDataFromDocument(document, filterValues, sheetName);
            }

            return dataTable;
        }
        
        public DataTable ObtainTable(Stream stream,  IEnumerable<FilterValue> filterValues, bool isEditable, bool disposeStreamAfterReading = true, string sheetName = "")
        {

            var dataTable = new DataTable();
            
            using (var document = SpreadsheetDocument.Open(stream, isEditable))
            {
                dataTable.DataRows = ExtractDataFromDocument(document, filterValues, sheetName);
            }
            
            return dataTable;
        }
        
        public async Task<DataTable> ObtainTableAsync(Stream stream,  IEnumerable<FilterValue> filterValues, bool isEditable, string sheetName = "")
        {

            var dataTable = new DataTable();
            
            using (var document = SpreadsheetDocument.Open(stream, isEditable))
            {
                var rows =  await ExtractDataFromDocumentAsync(document, filterValues, sheetName);
                dataTable.DataRows = rows.ToList();
            }
            
            return dataTable;
        }



        private List<DataRow> ExtractDataFromDocument(SpreadsheetDocument document,  IEnumerable<FilterValue> filterValues, string sheetName = "")
        {
            WorkbookPart workbookPart = document.WorkbookPart;
            SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
            SharedStringTable sst = sstpart.SharedStringTable;

            var worksheetPart = GetWorksheetPartByName(document, sheetName);
            
            if(worksheetPart == null)
                worksheetPart = workbookPart.WorksheetParts.Last();
            
            Worksheet sheet = worksheetPart.Worksheet;
                
            var rows = sheet
                .Descendants<Row>()
                .ToList();
                
            // найти заголовок согласно фильтру

            var rowMap = GetHeaderMap(sst, rows, filterValues);

            if (!rowMap.CellMaps.Any())
                return new List<DataRow>();

            // найти тело до первой пустой строки
                
            var bodyRows = ExtractBodyRows(sst, rowMap, rows);
            
            return bodyRows;
        }
        
        private async Task<IEnumerable<DataRow>> ExtractDataFromDocumentAsync(SpreadsheetDocument document,  IEnumerable<FilterValue> filterValues, string sheetName = "")
        {
            WorkbookPart workbookPart = document.WorkbookPart;
            SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
            SharedStringTable sst = sstpart.SharedStringTable;

            var worksheetPart = GetWorksheetPartByName(document, sheetName);
            
            if(worksheetPart == null)
                worksheetPart = workbookPart.WorksheetParts.Last();
            
            Worksheet sheet = worksheetPart.Worksheet;
                
            var rows = sheet
                .Descendants<Row>()
                .ToList();
                
            // найти заголовок согласно фильтру

            var rowMap = GetHeaderMap(sst, rows, filterValues);

            if (!rowMap.CellMaps.Any())
                return new List<DataRow>();
            
            var bodyRows = await ExtractBodyRowsAsync(sst, rowMap, rows);

            return bodyRows;
        }

        private List<DataRow> ExtractBodyRows(SharedStringTable sst, RowMap rowMap, IEnumerable<Row> rows)
        {
            var firstBodyRowIndex = rowMap.RowIndex + 1;
            var rowList = rows.Where(r => r.RowIndex >= firstBodyRowIndex).ToList();

            var  dataRows = new List<DataRow>();

            var index = firstBodyRowIndex;
            

            foreach (var row in rowList)
            {
                // если расстояние между непустыми строками больше 1, значит таблица закончилась
                if(row.RowIndex - index > 1)
                    break;
                
                var cells = ExtractCells(sst, row, rowMap.CellMaps);
                
                // если не получил заполненных ячеек, то таблица закончилась
                if(!cells.Any()) break;
                
                dataRows.Add(new DataRow
                {
                    DataCells = cells
                });

                index++;
            }

            return dataRows;
        }
        
        
        
        private async Task<IEnumerable<DataRow>> ExtractBodyRowsAsync(SharedStringTable sst, RowMap rowMap, IEnumerable<Row> rows)
        {
            var firstBodyRowIndex = rowMap.RowIndex + 1;
            var rowList = rows.Where(r => r.RowIndex >= firstBodyRowIndex).ToList();
            
            var extractRowsTasks = rowList.Select(row =>
                Task.Run(() =>
                {
                    var cells = ExtractCells(sst, row, rowMap.CellMaps);
                    
                    return new DataRow
                    {
                        RowIndex = row.RowIndex,
                        DataCells = cells
                    };
                    
                }));


            var dataRows = await Task.WhenAll(extractRowsTasks);
            
            dataRows =
                dataRows
                    .Where(r => r.DataCells.Any())
                    .OrderBy(r => r.RowIndex)
                    .ToArray();
            
            return dataRows;
        }
        

        private List<DataCell> ExtractCells(SharedStringTable sst, Row row, IEnumerable<CellMap> cellMaps)
        {

            var result = new List<DataCell>();

            var cellsByRow = GetCells(row);
            
            var cells = cellsByRow.Select(cell 
                =>   new {
                        value = GetCellValue(sst, cell),
                        column = GetColumnIndex(cell.CellReference)
                    }).ToList();

            // если все ячейки пустые, таблица либо пустая, либо закончилась
            if (cells.All(cell => string.IsNullOrEmpty(cell.value)))
                return result;

            foreach (var cellMap in cellMaps)
            {
                var cellByColumn = cells.FirstOrDefault(cell => cell.column == cellMap.ColumnIndex);
                
                result.Add(new DataCell
                {
                    Value = cellByColumn?.value ?? string.Empty,
                    PropertyInfo = cellMap.FilterValue.PropertyInfo
                });
            }

            return result;
        }
        
    }
}