using DataHandler.Excel.Implementation;
using System;
using System.Collections.Generic;
using System.Text;

namespace DataHandler.Tests
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
