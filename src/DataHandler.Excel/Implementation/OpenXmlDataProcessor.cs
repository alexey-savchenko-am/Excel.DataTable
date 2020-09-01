using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DataHandler.Excel.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DataHandler.Excel.Implementation
{
    public abstract class OpenXmlDataProcessor
    {
        
        protected uint InsertBorder(WorkbookPart workbookPart, Border border)
        {
            Borders borders = workbookPart.WorkbookStylesPart.Stylesheet.Elements<Borders>().First();
            borders.Append(border);
            return (uint)borders.Count++;
        }

        
        protected Border GenerateBorder()
        { 
            Border border2 = new Border();

            LeftBorder leftBorder2 = new LeftBorder(){ Style = BorderStyleValues.Thin };
            Color color1 = new Color(){ Indexed = (UInt32Value)64U };

            leftBorder2.Append(color1);

            RightBorder rightBorder2 = new RightBorder(){ Style = BorderStyleValues.Thin };
            Color color2 = new Color(){ Indexed = (UInt32Value)64U };

            rightBorder2.Append(color2);

            TopBorder topBorder2 = new TopBorder(){ Style = BorderStyleValues.Thin };
            Color color3 = new Color(){ Indexed = (UInt32Value)64U };

            topBorder2.Append(color3);

            BottomBorder bottomBorder2 = new BottomBorder(){ Style = BorderStyleValues.Thin };
            Color color4 = new Color(){ Indexed = (UInt32Value)64U };

            bottomBorder2.Append(color4);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            return border2;
        }
        
        protected uint InsertCellFormat(WorkbookPart workbookPart, CellFormat cellFormat)
        {
            CellFormats cellFormats = workbookPart.WorkbookStylesPart.Stylesheet.Elements<CellFormats>().First();
            cellFormats.Append(cellFormat);
            return (uint)cellFormats.Count++;
        }
        
        protected RowMap GetHeaderMap(SharedStringTable sst, IEnumerable<Row> rows, IEnumerable<FilterValue> filterValues)
        {
            var rowMap = new RowMap();
            
            foreach (Row row in rows)
            {
                var cellsMap = GetHeaderCellsMap(sst, row, filterValues);

                if (cellsMap == null || !cellsMap.Any()) continue;

                rowMap.RowIndex = Int32.Parse(row.RowIndex);
                rowMap.CellMaps = cellsMap;
                
                break;
            }

            return rowMap;
        }


        private List<CellMap> GetHeaderCellsMap(SharedStringTable sst, Row row, IEnumerable<FilterValue> filterValues)
        {

            var cells = row.Elements<Cell>().ToList();

            if (!cells.Any()) return null;
            
            var cellMaps = new List<CellMap>();
            
            string cellValue = null;

            int cellIndex = 0;
            
            foreach (Cell c in cells)
            {
                cellValue = GetCellValue(sst, c).Trim();

                if(cellValue == string.Empty) continue;
                
                var filterValue = filterValues.FirstOrDefault(fv => fv.Label == cellValue);

                if (filterValue == null) return null;
                
                cellMaps.Add(new CellMap
                {
                    ColumnIndex = GetColumnIndex(c.CellReference).Value,
                    FilterValue = filterValue
                });

                cellIndex++;
            }
            
            
            return cellMaps;
        }


        protected string GetCellValue(SharedStringTable sst, Cell cell)
        {
            if (cell.CellValue == null)
                return string.Empty;
            
            if ((cell.DataType != null) && (cell.DataType ==  CellValues.SharedString))
            {
                int ssid = int.Parse(cell.CellValue.Text);
                string str = sst.ChildElements[ssid].InnerText;

                return str.Trim();
            }

            return cell.CellValue.Text.Trim();
        }
        
        
        protected string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            } 

            return columnName;
        }
        
        protected int? GetColumnIndex(string cellReference)
        {
            if (string.IsNullOrEmpty(cellReference))
            {
                return null;
            }

            //remove digits
            string columnReference = Regex.Replace(cellReference.ToUpper(), @"[\d]", string.Empty);

            int columnNumber = -1;
            int mulitplier = 1;

            //working from the end of the letters take the ASCII code less 64 (so A = 1, B =2...etc)
            //then multiply that number by our multiplier (which starts at 1)
            //multiply our multiplier by 26 as there are 26 letters
            foreach (char c in columnReference.ToCharArray().Reverse())
            {
                columnNumber += mulitplier * ((int)c - 64);

                mulitplier = mulitplier * 26;
            }
            

            //the result is zero based so return columnnumber + 1 for a 1 based answer
            //this will match Excel's COLUMN function
            return columnNumber + 1;
        }

        
        protected  WorksheetPart  GetWorksheetPartByName(SpreadsheetDocument document, 
                string sheetName)
        {
            IEnumerable<Sheet> sheets =
                document.WorkbookPart.Workbook.GetFirstChild<Sheets>().
                    Elements<Sheet>().Where(s => s.Name == sheetName);

            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.

                return null;
            }

            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)
                document.WorkbookPart.GetPartById(relationshipId);
            return worksheetPart;

        }
        
          protected  Cell GetCell(Worksheet worksheet, 
            string columnName, uint rowIndex)
        {
            Row row = GetRow(worksheet, rowIndex);

            if (row == null)
                return null;

            return row.Elements<Cell>()
                .Where(c => string.Compare(c.CellReference.Value, columnName +  rowIndex, true) == 0).First();
        }
        
        private Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            return worksheet.GetFirstChild<SheetData>().
                Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        } 
        
       
        
        protected  void InsertSharedStringItem(SpreadsheetDocument ssDoc, Cell cell, string valueText)
        {
            var num = InsertSharedStringItem(ssDoc, valueText);
            cell.CellValue = new CellValue(num.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
        }
        
        protected  int InsertSharedStringItem(SpreadsheetDocument ssDoc, string text)
        {
            SharedStringTablePart part;
            if (ssDoc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Any())
            {
                part = ssDoc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                part = ssDoc.WorkbookPart.AddNewPart<SharedStringTablePart>("rId6");
                part.SharedStringTable = new SharedStringTable() { Count = 1, UniqueCount = 1 };
            }
            int num = 0;
            foreach (SharedStringItem item in part.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return num;
                }
                num++;
            }
            part.SharedStringTable.AppendChild<SharedStringItem>(new SharedStringItem(new OpenXmlElement[] { new Text(text) }));
            part.SharedStringTable.Save();
            return num;
        }
        
        
        
        protected List<Cell> GetCells(Row row)
            => row.Elements<Cell>().ToList();


       protected struct RowMap
       {
           public int RowIndex { get; set; }
           
           public List<CellMap> CellMaps { get; set; }
       }

       protected struct CellMap
       {
           public int ColumnIndex { get; set; }

           public FilterValue FilterValue { get; set; }
       }
    }
}