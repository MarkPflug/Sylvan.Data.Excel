# Sylvan.Data.Excel

A cross-platform .NET library for reading Excel data files in .xlsx, .xlsb and .xls formats.
Provides readonly, row by row, forward-only access to the data.
There is no support for creating or editing Excel files.
Provides a familiar API via `DbDataReader`, which is ideal for accessing rectangular, tabular data sets.
The library is a purely managed implementation with no external dependencies.

This library is currently the [fastest and lowest allocating](https://github.com/MarkPflug/Benchmarks/blob/main/docs/ExcelBenchmarks.md) 
library for reading Excel data files in the .NET ecosystem, for all supported formats.

## Installing

This library is still relatively immature and you might encounter issues while using it.
If you do encounter any bugs, please report an issue in the github repository.
Be aware that I will be unlikely to investigate any issue unless an example file can be provided reproducing the issue.

[Sylvan.Data.Excel Nuget Package](https://www.nuget.org/packages/Sylvan.Data.Excel/)

`Install-Package Sylvan.Data.Excel`

## Basic Usage
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

Exporting Excel data to CSV(s): (using Sylvan.Data.Excel and Sylvan.Data.Csv)
```C#
using Sylvan.Data.Excel;
using Sylvan.Data.Csv;

using var edr = ExcelDataReader.Create("data.xls");

do 
{
	var sheetName = edr.WorksheetName;
	using CsvDataWriter cdw = CsvDataWriter.Create("data-" + sheetName + ".csv")
	cdw.Write(edr);
} while(edr.NextResult());
```
