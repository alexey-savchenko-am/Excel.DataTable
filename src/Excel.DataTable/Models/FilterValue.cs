using System;
using System.Collections.Generic;
using System.Reflection;

namespace Excel.DataTable.Models
{
    public class FilterValue
    {
        public PropertyInfo PropertyInfo{ get; set; }
        
        public string Label { get; set; }

        public Type Type { get; set; }
        
    }

    public class FilterSet
    {
        public List<FilterValue> FilterValues { get; set; }
            = new List<FilterValue>();
    }
    
}