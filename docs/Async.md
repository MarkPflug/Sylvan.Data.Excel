# Async Support

ExcelDataReader and ExcelDataWriter both support "async" operations.
In order to ensure proper async behavior, you must use the `CreateAsync`
method to create instances. In this mode, the entire file will be 
asynchronously buffered into memory.

Reading:
```
// this line will buffer the entire file into memory.
await using var edr = ExcelDataReader.CreateAsync("jumbo.xlsx");

while(await edr.ReadAsync())
{
  // ...
}

```

Writing:
```
await using var edw = ExcelDataWriter.CreateAsync("jumbo.xlsx");

edw.WriteAsync(myDataReader, "MyData");
edw.WriteAsync(myOtherDataReader, "MoreData");

// when the ExcelDataWriter is asynchronously disposed
// the buffered file is asynchronously written to the output.
```