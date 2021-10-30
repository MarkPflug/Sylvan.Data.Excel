# Sylvan.Data.Excel

The cross-platform .NET library for reading Excel data files in .xls and .xlsx format.
Provides readonly, row by row, forward-only access to the data.
There is no support for creating or editing Excel files.
Provides a familiar API via `DbDataReader`, which is ideal for accessing rectangular, tabular data sets.

This library is currently the [fastest and lowest allocating](https://github.com/MarkPflug/Benchmarks/blob/main/docs/ExcelBenchmarks.md) library for reading Excel data files in the .NET ecosystem.

## Installing

This library is still in early development and you might encounter issues while using it.
If you do find bugs, please report an issue. However, be aware that I will be unlikely to be able 
to investigate any issue unless a file reproducing the problem is provided with the issue.

[Sylvan.Data.Excel Nuget Package](https://www.nuget.org/packages/Sylvan.Data.Excel/)

`Install-Package Sylvan.Data.Excel`

## Basic Usage
```C#

// ExcelDataReader derives from System.Data.DbDataReader
using ExcelDataReader edr = ExcelDataReader.Create("data.xls");

// Same API for both xls and xlsx files.
// edr = ExcelDataReader.Create("data.xlsx");

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
	}
	// iterates sheets
} while(edr.NextResult());

```

Exporting Excel data to CSV(s): (using Sylvan.Data.Excel and Sylvan.Data.Csv)
```C#

using var edr = ExcelDataReader.Create("data.xls");

do 
{
	var sheetName = edr.WorksheetName;
	using CsvDataWriter cdw = CsvDataWriter.Create("data-" + sheetName + ".csv")
	cdw.Write(edr);
} while(edr.NextResult());

```
