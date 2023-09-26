# Schema Information

Excel is a schemaless file format; a single column can contain any mix of data types.
The `DbDataReader`, which `ExcelDataReader` extends, exposes APIs for inspecting the schema of the data.
Since there is no guaranteed uniformity to the data in Excel, 
the default schema of `ExcelDataReader` exposes all columns as nullable strings. 
This means that all calls to `GetFieldType` will return `typeof(string)`, for all columns.
GetColumnSchema and GetSchemaTable will likewise indicate that all columns are strings.

Calling `GetString` will always return non-null string representation of the value in the cell.
The string representation will not always be the exact format as displayed in Excel, however.
For example, date values will always be returned as an ISO8601 formatted string. 
Numeric values might also be returned in a different number format than displayed in Excel, 
but the numerical value will represent the actual value in Excel.

ExcelDataReader exposes a method `GetExcelDataType(int ordinal)`, which can be used to determine
the data type for a particular cell; one of the interal Excel types: null, numeric, datetime, string, boolean or error.
This method can return different types for different rows in the same column, where `GetFieldType`
will always return the same type for all rows.
Any of the strongly-typed accessors (GetString, GetInt32, GetDouble) can be used to access a cell, 
and will return a value if the internal Excel value can be safely converted.

# Strongly-Typed Schema

It is often useful to apply a strict schema to the data in an Excel file when the columns are known
to be constrained to values of a uniform type. This allows the data to be read in a strongly-typed way
when loading into a `DataTable` or being processed by certain schema-aware tools like `SqlBulkCopy`.

Sylvan.Data.Excel allows applying a schema to Excel data via the `IExcelSchemaProvider` interface.
This interface allows providing a different schema per-sheet in a workbook. 

The `ExcelSchema` class provides a concrete implementation of `IExcelSchemaProvider`, and is the easiet
way to apply schema information to spreadsheet data.

As an example, given the data in the following table:

| Id | Name  | Value | ReleaseDate | Notes |
|----|-------|-------|-------------|-------|
| 1  | Alpha |  7.50 | 2020-03-01  | Discontinued |
| 2  | Beta  | 12.75 | 2022-01-05  | |

The following is an example of applying a schema to this data.
It uses the `Sylvan.Data.Schema` type (which comes from the Sylvan.Data nuget package) 
to create a strongly-typed schema definition from a string. 
This results in `table` being loaded with strongly-typed values, instead of strings.

```CSharp
using Sylvan.Data;
using Sylvan.Data.Excel;
using System.Data;

// uses Sylvan.Data package to create a schema definition.
var schema = Schema.Parse("Id:int,Name:string,Value:decimal,ReleaseDate:date,Notes:string?");

// creates an Excel schema that can apply the above schema to an Excel worksheet.
var excelSchema = new ExcelSchema(hasHeaders: true, schema);

using (var data = ExcelDataReader.Create("data.xlsx", new ExcelDataReaderOptions { Schema = excelSchema })) 
{
    var table = new DataTable(data.WorksheetName);
    table.Load(data);
    // table will be loaded with strongly-typed values.
}
```