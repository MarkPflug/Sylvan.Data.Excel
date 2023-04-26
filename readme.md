# Sylvan.Data.Excel

A cross-platform .NET library for reading and writing Excel data files. The most commonly used formats: .xlsx, .xlsb and .xls, are supported for reading, while .xlsx and .xlsb formats are supported for writing.

ExcelDataReader provides readonly, row by row, forward-only access to the data. It provides a familiar API via `DbDataReader`, which is ideal for accessing rectangular, tabular data sets. It exposes a single, unified API for accessing all supported file formats.

ExcelDataWriter supports writing data from a `DbDataReader` to an Excel worksheet in .xlsx and .xlsb formats. ExcelDataWriter can only write a new Excel, it cannot be used to edit or append to existing files. It does not support custom formatting, charts, or other common features. It is meant only to export raw, flat data to Excel.

The library is a purely managed implementation with no external dependencies.

This library is currently the [fastest and lowest allocating](https://github.com/MarkPflug/Benchmarks/blob/main/docs/ExcelReaderBenchmarks.md) 
library for reading Excel data files in the .NET ecosystem, for all supported formats.

If you encounter any issues while using this library, please report an issue in the github repository.
Be aware that I will be unlikely to investigate any issue unless an example file can be provided reproducing the issue.

## Installing

[Sylvan.Data.Excel Nuget Package](https://www.nuget.org/packages/Sylvan.Data.Excel/)

`Install-Package Sylvan.Data.Excel`

## ExcelDataReader

ExcelDataReader derives from DbDataReader, so it exposes an API that should be familiar for anyone who has worked with ADO.NET before. The field accessors allow reading data in an efficient, strongly-typed manner: `GetString`, `GetInt32`, `GetDateTime`, `GetBoolean`, etc. 

The `GetExcelDataType` method allows inspecting the native Excel data type of a cell, which may vary from row to row. `FieldCount`, returns the number of columns in the header row, and doesn't change while processing each row in a sheet. `RowFieldCount` returns the number of fields in the current row, which might vary from row to row, and can be used to access cells in a "jagged", non-rectangular file.

### Reading Raw Data

The ExcelDataReader provides a forward only, row by row access to the data in a worksheet. It allows iterating over sheets using the `NextResult()` method, and iterating over rows using the `Read()` method. Fields are accessed using standard accessors, most commonly `GetString()`. `GetString()` is designed to not throw an exception, except in the case that a cell contains a formula error.

```C#
using Sylvan.Data.Excel;

// ExcelDataReader derives from System.Data.DbDataReader
// The Create method can open .xls, .xlsx or .xlsb files.
using ExcelDataReader edr = ExcelDataReader.Create("data.xls");

do 
{
	var sheetName = edr.WorksheetName;
	// enumerate rows in current sheet.
	while(edr.Read())
	{
		// iterate cells in row.
		// can use edr.RowFieldCount when sheet contains jagged, non-rectangular data
		for(int i = 0; i < edr.FieldCount; i++)
		{
			var value = edr.GetString(i);
		}
		// Can use other strongly-typed accessors
		// bool flag = edr.GetBoolean(0);
		// DateTime date = edr.GetDateTime(1);
		// decimal amt = edr.GetDecimal(2);
	}
	// iterates sheets
} while(edr.NextResult());
```

### Bind Excel data to objects using Sylvan.Data

The Sylvan.Data library includes a general-purpose data binder that can bind a DbDataReader to objects.
This can be used to easily read an Excel file as a series of strongly typed objects.

```C#
using Sylvan.Data;
using Sylvan.Data.Excel;

using var edr = ExcelDataReader.Create("data.xlsx");
foreach (MyRecord item in edr.GetRecords<MyRecord>())
{
    Console.WriteLine($"{item.Name} {item.Quantity}");
}

class MyRecord
{
    public int Id { get; set; }
    public string Name { get; set; }
    public int? Quantity { get; set; }
    public DateTime Date { get; set; }
}
```

### Converting Excel data to CSV(s) 

The Sylvan.Data.Csv library can be used to convert Excel worksheet data to CSV.

```C#
using Sylvan.Data.Excel;
using Sylvan.Data.Csv;

using var edr = ExcelDataReader.Create("data.xlsx");

do 
{
	var sheetName = edr.WorksheetName;
	using CsvDataWriter cdw = CsvDataWriter.Create("data-" + sheetName + ".csv")
	cdw.Write(edr);
} while(edr.NextResult());
```

## ExcelDataWriter

The `ExcelDataWriter` type is used to create Excel workbooks and write `DbDataReader` data as worksheets.

```C#
// *critical* to dispose (using) ExcelDataWriter.
using var edw = ExcelDataWriter.Create("data.xlsx");
DbDataReader dr;
dr = GetQueryResults("UserReport");
edw.Write(dr, "UserReport");
dr = GetQueryResults("SecurityAudit");
edw.Write(dr, "SecurityAudit");
```
