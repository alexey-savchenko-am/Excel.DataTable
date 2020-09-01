using System;
using System.Collections.Generic;
using System.Linq;
using AutoFixture;
using DataHandler.Excel;
using DataHandler.Excel.Implementation;
using DataHandler.Excel.Models;
using Xunit;

namespace DataHandler.Tests.UnitTests
{
    public class ExcelDataHandlerTests
    {
        [Fact]
        public void ExtractingDataFromExcelFileSuccess()
        {
            var data =
                new ExcelDataParser<SalesOrdersDataModel>(
                        new OpenXmlDataObtainer(), 
                        new OpenXmlDataWriter())
                    .Bind("./SampleData.xlsx")
                    .ExtractData(true, "SalesOrders")
                    .Result;
            
            Assert.True(data.Any());
        }
        
        [Fact]
        public void WritingDataFromExcelFileSuccess()
        {
            var fixture = new Fixture();

            var testRecords = 
                fixture
                    .CreateMany<SalesOrdersDataModel>()
                    .ToList();

            IDataParser<SalesOrdersDataModel> parser = null;
            try
            {
                parser =
                    new ExcelDataParser<SalesOrdersDataModel>(
                            new OpenXmlDataObtainer(), 
                            new OpenXmlDataWriter())
                        .Bind("./SampleData.xlsx", true)
                        .WriteData(testRecords, RowStyles.Simple, true, "SalesOrders");
            }
            finally
            {
                parser.Clear();
            }
   

            var data = parser
                .Bind("./SampleData.xlsx")
                .ExtractData(true, "SalesOrders")
                .Result;

            var orderEqualityComparer = new OrderEqualityComparer();
            var result = testRecords.All(x => data.Contains(x, orderEqualityComparer));


            Assert.True(result);
        }



       
    }
}