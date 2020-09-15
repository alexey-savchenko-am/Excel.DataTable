using System.Collections.Generic;

namespace Excel.DataTable.Tests
{
    internal class OrderEqualityComparer
        : IEqualityComparer<SalesOrdersDataModel>
    {
        public bool Equals(SalesOrdersDataModel x, SalesOrdersDataModel y)
        {
            return
                x.ItemName == y.ItemName
                && x.Price == y.Price
                && x.Region == y.Region
                && x.Units == y.Units
                && x.CustomerName == y.CustomerName
                && x.TotalPrice == y.TotalPrice
                && x.OrderDate == y.OrderDate;
        }

        public int GetHashCode(SalesOrdersDataModel obj)
        {
            return obj.GetHashCode();
        }
    }
}