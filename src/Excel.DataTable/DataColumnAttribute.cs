using System;

namespace Excel.DataTable
{
    public class DataColumnAttribute 
        : Attribute
    {
        public DataColumnAttribute(string name)
        {
            this.ColumnName = name;
        }

        public string ColumnName { get; set; }
    }
}