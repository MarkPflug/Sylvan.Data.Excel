# Sylvan.Data.Excel Release Notes
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
