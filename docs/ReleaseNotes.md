# Sylvan.Data.Excel Release Notes

_0.4.20_
- Fixes a bug where advancing to the next worksheet can result in the incorrect number of columns being reported when reading xlsx files.

_0.4.19_
- Fix a bug where headers might not be read properly on sheets after the first sheet.

_0.4.18_
- Fix a bug where column filters were applied to one too many columns when writing .xlsb files.
- Fix a bug where data would be skipped when headers were disabled and a worksheet started with blank line(s).

_0.4.17_
- Exclude phonetic component when reading string values.
- Allow invalid ref values, which Excel appears to treat as missing. 
- SharedString tables are read lazily instead of eagerly for .xlsx and .xlsb files, allowing
   faster access to early records in some cases.

_0.4.16_
- Adds ExcelFileType class that exposes constants about supported Excel formats: extensions and content types.
- Add `Obsolete` to `TryOpenWorksheetAsync`, use `TryOpenWorksheet` instead.
- Add `CreateAsync` to `ExcelDataReader` and `ExcelDataWriter` to enable async operation (not supported on .NET Framework versions).

_0.4.15_
- Fix a bug that prevented .xlsx reader from working on .NET Framework versions.
- Fix a bug where FieldRowCount would be incorrect on empty rows.

_0.4.14_
- Add `CompressionLevel` as a configuration option for ExcelDataWriter.
    Changes the default compression level to `Fastest`, which produces 
    *slightly* larger files, but takes significantly less time.

_0.4.13_
- Fix for .xlsx/.xlsb files that contain OpenPackage relations with rooted paths.
   Third party libraries sometimes write files in this way.

_0.4.12_
- Adds support for Excel files using 1904 (mac?) date mode.

_0.4.11_
 - Adds support for writing .xlsb files.
 - Small optimization for .xlsx reading in some scenarios.
 - Add support for writing enum typed columns.

_0.4.10_
 - Fixes a bug where calling `ExcelDataReader.Read` again after it has already returned `false` could 
   result in the reader spinning in a tight loop.

_0.4.9_
 - Reading of .xlsx and .xlsb files is fixed to correctly use open packaging specifications.
   This should allow more third-party sourced files to be read.

_0.4.8_
 - Fix an issue with reading .xlsb files where moving to the next sheet didn't reset the reader state,
   and could cause issues reading certain files.

_0.4.7_
 - Xlsx files can be read that contain malfored cell references. 
    When opened in Excel these files present no error, and the cell defaults to what would be 
    the next cell in the sheet, as if no cell reference were present at all.

_0.4.6_
  - Allow enumerating and opening specific worksheets.
  - Add the ability to read Excel data dynamically (`ExcelSchema.Dynamic`), where each cell
    value is dynamically determined when accessed via `ExcelDataReader.GetValue(int)`.

_0.4.5_
  - Fix a bug where underscore characters where escaped when they didn't need to be. 
    Certain versions of Excel could complain about this and incorrectly repair the over-escaping.

_0.4.4_
  - Fix a bug in handling OpenXml character escaping that could cause certain strings in.xlsx files
   to be read incorrectly.

_0.4.3_
  - Fix .xlsx string handling for rare whitespace in shared string table.
  - Fix .xlsx whitespace handling to correctly preserve whitespace when reading and writing.
  - Xlsx writer now throws an exception when a string exceeds the maximum length allowed by Excel (32k).
      Use `ExcelDataWriterOptions.TruncateStrings = true` to truncate such strings when writing.

_0.4.2_
  - Fix ExcelDataWriter formatting to use InvariantCulture to provide consistent behavior in all cultures.

_0.4.1_
  - Fix version metadata format written by ExcelDataWriter which was causing warnings on some machines.

_0.4.0_
  - Adds ExcelDataWriter which supports writing a `DbDataReader` to an .xlsx worksheet.
  - Fixes the encoding of certain characters.

_0.3.4_
  - Fix .xls reader to not require files use Row/DBCell records, which are not written by all sources.
  - Fix RowNumber being incorrect when reading with certain configurations.

_0.3.3_
  - Adds ExcelSchema support for renaming columns via BaseColumnName.

_0.3.2_
  - Fix for format not being applied to some cells (Crystal Decisions).
  - Add DateTimeFormat option, which applies when dates are stored as strings.
  - Replace FormatException with InvalidCastException, as documented by DbDataReader.

_0.3.1_
  - Fix reading .xlsx files that don't specify shared string count (Crystal Decisions).
  - Skipping sheets (`NextResult`) in .xls files is faster.

_0.3.0_
  - Add non-allocating implementation of `GetFieldValue<T>()`.
  - Support for reading enum values via `GetFieldValue<T>()`.
  - Add support for `GetChar(int ordinal)`.
  - Add `IExcelSchemaProvider.GetFieldCount(ExcelDataReader)` to allow the provider to explicitly
    define the number of columns. Previously, the number of columns in the current row would be used.
  - Fix an issue where calling `Initialize()` on the first row would not work as expected.

_0.2.3_
  - Fix an issue where a trailing row appears empty, but contains a cell with an empty string value. 
    Previously, calls to `Read()` would return `true` and there rows would be processed. 
    With this fix, these rows will be skipped `Read()` will return `false`.

_0.2.2_
  - Add the ability to convert string to boolean in GetBoolean via `ExcelDataReaderOptions.True/FalseString`.
  - Unify implementation of `GetOrdinal(string)` to use case-insensitive matching.

_0.2.1_
- Fix reading .xlsx files created by JasperReports.
  - No `count` on xfCells elements.
  - Empty `inlineStr` values.

_0.2.0_
- Remove InitializeSchema API and replace with Initialize.
- Fix for exception on unknown format, now defaults to .NET default double format.
- Fix for schema mappings falling back to ordinal not working.
- Fix for empty shared string values created by AG Grid.
- Fix handling of certain formats that would be misidentified as date-kind formats.
- Fix Date formatting bug on netstandard2.0 where ticks values wouldn't be formatted as null characters.
- Fix for reading formulas that produce empty values.

_0.1.12_
- Fix for reading .xlsx files with inline-string values, believed to be created with Apache POI.

_0.1.11_
- Fix bug with reading certain numeric values in xlsb files.
- Fix bug where stream wasn't closed when reader was disposed.
- Fix `RowCount` for .xlsx and .xlsb files.
- GetName returns empty string for out of range access, instead of throwing.

_0.1.10_
- Fix bug with .xlsx file containing empty element in shared string table.

_0.1.9_
- Fix bug with .xlsx shared strings table containing empty string element.
- Fix buffer management bug with .xls files.

_0.1.8_
- Various bug fixes.

_0.1.7_
- Add .xslb support.
- Option to process hidden sheets.

_0.1.5_
- Add `RowFieldCount` property to determine number of fields in the current row.

_0.1.4_
- Fix behavior of GetValue to honor the data type of the schema instead of the Excel type of the column.
- Implement `GetSchemaTable()` so DataTable.Load functions correctly.
- Skip rows containing empty data in xlsx files.
- Fix a NotImplementedException in netstandard 2.0.
