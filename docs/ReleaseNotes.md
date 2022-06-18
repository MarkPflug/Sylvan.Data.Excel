# Sylvan.Data.Excel Release Notes
_0.2.0_
- Remove InitializeSchema API and replace with Initialize.
- Fix for exception on unknown format, now defaults to .NET default double format.
- Fix for schema mappings falling back to ordinal not working.

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
- Fix behavior of GetValue to honor the data type of the schema instead of the excel type of the column.
- Implement `GetSchemaTable()` so DataTable.Load functions correctly.
- Skip rows containing empty data in xlsx files.
- Fix a NotImplementedException in netstandard 2.0.
