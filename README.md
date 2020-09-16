[![NuGet version (Excel.DataTable)](https://img.shields.io/nuget/v/IO.Pipeline.svg?style=flat-square&color=blue)](https://www.nuget.org/packages/Excel.DataTable)

# Excel.DataTable
Allows to extract or write data easily from/to Excel tables

# Usage

The tool allows to parse excel file and retrieve or write some data to excel table.

Assume we have an excel file which contains table like this one:

![table sample](https://github.com/goOrn/DataHandler/blob/master/screenshots/table.JPG?raw=true)

First of all, we should create a model, each property of which contains information about physical columns of the excel table.
DataColumn attribute should have a name of physical column of the table, but the property itself can have an arbitrary name:

```
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
```

## ExcelDataParser

Use ExcelDataParser to read or write data.
You should specify generic type as SalesOrdersDataModel.

```
  var dataParser =
     new ExcelDataParser<SalesOrdersDataModel>(
          new OpenXmlDataObtainer(), 
          new OpenXmlDataWriter());
```

## Bind

Bind data parser with physical excel file on disk or stream with the following command:

```
  dataParser.Bind("./SampleData.xlsx");
```

## Extract data

To extract data from file use the command ExtractData.
First parameter is boolean value to clear streams after reading file.
The secon one is a Excel sheet name, where the table specified:

```
  dataParser.ExtractData("SalesOrders")
```

## Result 

Use property Result to get data as a list of objects with type SalesOrdersDataModel.

Full code of extractig data from excel file should look like this one:

```
  var data =
    new ExcelDataParser<SalesOrdersDataModel>(
          new OpenXmlDataObtainer(), 
          new OpenXmlDataWriter())
       .Bind("./SampleData.xlsx")
       .ExtractData("SalesOrders")
       .Result;
```

## DataParser

Use DataParser class to use default OpenXmlDataObtainer and OpenXmlDataWriter.
ExcelDataParser implements IDisposable interface to clear streams after reading or writing data.
So you need to use parser within using block:

```
var result = new List<SalesOrdersDataModel>();
using (var dataParser = new DataParser<SalesOrdersDataModel>())
{
    result = 
         dataParser
           .Bind("./SampleData.xlsx")
           .ExtractData("SalesOrders")
           .Result;
}
```

## Write data

You are able to write data to excel table:

```
 var fixture = new Fixture();
 
var testRecords = 
    fixture
    .CreateMany<SalesOrdersDataModel>()
    .ToList();
    
using (var dataParser = new DataParser<SalesOrdersDataModel>())
{
    result = 
         dataParser
           .Bind("./SampleData.xlsx")
           .WriteData(testRecords, RowStyles.Simple, false, "SalesOrders");
}

```
The first parameter of WriteData method is a list of records of type SalesOrdersDataModel.
The second one is a member of RowStyle enum.
If you set the third parameter to true, the stream will not be cleared after data writing.
The last parameter is SheetName where the table exists.

