using System;

namespace DataHandler.Excel
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