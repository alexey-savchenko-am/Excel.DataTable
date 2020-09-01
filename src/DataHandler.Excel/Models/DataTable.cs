using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace DataHandler.Excel.Models
{
    public class DataTable
    {
        public List<DataRow> DataRows { get; set; }
            = new List<DataRow>();
    }

    public static class DataTableExtensions
    {
        public static DataTable Combine (this DataTable firstTable, ref DataTable secondTable)
        {
            var dataRows = firstTable.DataRows.Concat(secondTable.DataRows);
            
            var result = new DataTable()
            {
                DataRows = dataRows.ToList()
            };

            secondTable = result;

            return result;
        }
        
        public static async Task<DataTable> CombineAsync (this DataTable firstTable, DataTable secondTable)
        {
            var dataRows = firstTable.DataRows.Concat(secondTable.DataRows);
            
            var result = new DataTable()
            {
                DataRows = dataRows.ToList()
            };

            secondTable = result;

            return result;
        }
    }


}