# Sylvan.Data.Excel

The .NET library for reading Excel data files in .xls and .xlsx format.
This provides read-only, row by row, forward-only access to the data.
There is no support for creating or editing Excel files.
Provides a familiar API via DbDataReader, which is ideal for accessing rectangular, columnar data sets.

## Installing

This library is still a pre-release version and you will likely encounter issues while using it.
If you do find bugs, please report an issue. However, be aware that I will be unlikely to be able 
to fix an issue unless an excel file exhibiting the problem is provided with the issue.

[Sylvan.Data.Excel Nuget Package](https://www.nuget.org/packages/Sylvan.Data.Excel/)

`Install-Package Sylvan.Data.Excel -IncludePrerelease`

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

## Tests

Many of the tests will currently fail because they depend on larger data files that I chose to
keep external to the repository. I intend to use a git sub-module or some other solution to keep them
separate from this repo.
