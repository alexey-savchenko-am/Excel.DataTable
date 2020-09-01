using System.Reflection;

namespace DataHandler.Excel.Models
{
    public class DataCell
    {
        public PropertyInfo PropertyInfo { get; set; }
        public string Value { get; set; }
    }
}