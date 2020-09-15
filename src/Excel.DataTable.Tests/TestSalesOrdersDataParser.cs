using Excel.DataTable.Implementation;

namespace Excel.DataTable.Tests
{
    public class TestSalesOrdersDataParser
        : ExcelDataParser<SalesOrdersDataModel>
    {
        public TestSalesOrdersDataParser()
          : base(new OpenXmlDataObtainer(), new OpenXmlDataWriter())
        {

        }
    }
}
