using System;
using System.Collections.Generic;
using System.Linq;
using AutoFixture;
using DataHandler.Excel.Models;
using Xunit;

namespace DataHandler.Tests.UnitTests
{
    public class ExcelDataHandlerTests
    {
        [Fact]
        public void ExtractingDataFromExcelFileSuccess()
        {
            var result = new List<SalesOrdersDataModel>();

            using (var parser = new TestSalesOrdersDataParser())
            {
                   result = 
                       parser
                        .Bind("./SampleData.xlsx")
                        .ExtractData("SalesOrders")
                        .Result;
            } 
       
            Assert.True(result.Any());
        }
        
        [Fact]
        public void WritingDataFromExcelFileSuccess()
        {
            var fixture = new Fixture();

            var testRecords = 
                fixture
                    .CreateMany<SalesOrdersDataModel>()
                    .ToList();

            var data = new List<SalesOrdersDataModel>();

            using (var parser = new TestSalesOrdersDataParser())
            {
                parser
                    .Bind("./SampleData.xlsx", true)
                    .WriteData(testRecords, RowStyles.Simple, false, "SalesOrders");

                data = parser
                  .Bind("./SampleData.xlsx")
                  .ExtractData("SalesOrders")
                  .Result;
            }

            var orderEqualityComparer = new OrderEqualityComparer();
            var result = testRecords.All(x => data.Contains(x, orderEqualityComparer));

            Assert.True(result);
        }



       
    }
}