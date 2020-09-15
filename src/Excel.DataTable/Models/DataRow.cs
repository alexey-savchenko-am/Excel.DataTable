using System;
using System.Collections.Generic;

namespace Excel.DataTable.Models
{
    public class DataRow
    {
        public UInt32 RowIndex { get; set; }
        public List<DataCell> DataCells { get; set; }
            = new List<DataCell>();
    }
}