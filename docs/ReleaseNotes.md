# Sylvan.Data.Excel Release Notes


_0.1.4_
- Fix behavior of GetValue to honor the data type of the schema instead of the excel type of the column.
- Implement `GetSchemaTable()` so DataTable.Load functions correctly.
- Skip rows containing empty data in xlsx files.
- Fix a NotImplementedException in netstandard 2.0.
