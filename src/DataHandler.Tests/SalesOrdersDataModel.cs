using DataHandler.Excel;

namespace DataHandler.Tests
{
    public class SalesOrdersDataModel
    {
        [DataColumn("OrderDate")]
        public string OrderDate { get; set; }
        
        [DataColumn("Region")]
        public string Region { get; set; }
        
        [DataColumn("Rep")]
        public string CustomerName { get; set; }
        
        [DataColumn("Item")]
        public string ItemName { get; set; }
        
        [DataColumn("Units")]
        public string Units { get; set; }
        
        [DataColumn("Unit Cost")]
        public string Price { get; set; }
        
        [DataColumn("Total")]
        public string TotalPrice { get; set; }
    }
}