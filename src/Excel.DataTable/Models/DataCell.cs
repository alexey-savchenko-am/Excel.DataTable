using System.Reflection;

namespace Excel.DataTable.Models
{
    public class DataCell
    {
        public PropertyInfo PropertyInfo { get; set; }
        public string Value { get; set; }
    }
}