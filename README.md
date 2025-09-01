# Excel.DataTable
A lightweight library to **extract data from Excel tables** into strongly-typed models or **write data back** into existing tables.

[![NuGet version (Excel.DataTable)](https://img.shields.io/nuget/v/Excel.DataTable.svg?style=flat-square&color=blue)](https://www.nuget.org/packages/Excel.DataTable)
[![Downloads](https://img.shields.io/nuget/dt/Excel.DataTable?style=flat-square&color=blue)]()

# Features
- Map Excel columns to C# model properties via attributes.  
- Extract data from Excel sheets as strongly-typed lists.  
- Write data into Excel tables with row formatting options.  
- Built on **OpenXML** (no Excel installation required).  
- Supports both file paths and streams.  

# Installation

```bash
dotnet add package Excel.DataTable
```
# Quick Start

```csharp
using (var parser = new DataParser<SalesOrdersDataModel>())
{
    // Read
    var data = parser
        .Bind("./SampleData.xlsx")
        .ExtractData("SalesOrders")
        .Result;

    // Write
    parser.WriteData(data, RowStyles.Simple, false, "SalesOrders");
}
```


# Usage

The tool allows you to parse an Excel file and either retrieve or write data to a table.

Assume we have an Excel file which contains a table like this:

![table sample](https://github.com/goOrn/DataHandler/blob/master/screenshots/table.JPG?raw=true)

First of all, you need to create a model where each property corresponds to a column of the Excel table.
The DataColumn attribute should contain the name of the actual Excel column, while the property name itself can be arbitrary:

```csharp
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
You should specify the generic type, e.g. SalesOrdersDataModel:

```csharp
  var dataParser =
     new ExcelDataParser<SalesOrdersDataModel>(
          new OpenXmlDataObtainer(), 
          new OpenXmlDataWriter());
```

## Bind

Bind the parser with a physical Excel file on disk or with a stream:

```csharp
  dataParser.Bind("./SampleData.xlsx");
```

## Extract data

To extract data from the file, use the ExtractData method.
Specify the sheet name where the table is located:

```csharp
  dataParser.ExtractData("SalesOrders")
```

## Result 

Use the Result property to get data as a list of objects of type SalesOrdersDataModel.

Full code of extracting data looks like this:

```csharp
  var data =
    new ExcelDataParser<SalesOrdersDataModel>(
          new OpenXmlDataObtainer(), 
          new OpenXmlDataWriter())
       .Bind("./SampleData.xlsx")
       .ExtractData("SalesOrders")
       .Result;
```

## DataParser

Use the DataParser class to simplify initialization with default OpenXmlDataObtainer and OpenXmlDataWriter.

ExcelDataParser implements IDisposable, so you should use it within a using block:

```csharp
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

You can also write data into an Excel table:

```csharp
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

