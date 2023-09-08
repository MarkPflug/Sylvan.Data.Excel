# Async Support

ExcelDataReader and ExcelDataWriter both support "async" operations.
In order to ensure proper async behavior, you must use the `CreateAsync`
method to create instances. In this mode, the entire file must be 
buffered in memory, and all IO will be handled asynchronously.

The CreateAsync methods are only supported on .NET Core versions.

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
// must use async disposal
await using var edw = ExcelDataWriter.CreateAsync("jumbo.xlsx");

edw.WriteAsync(myDataReader, "MyData");
edw.WriteAsync(myOtherDataReader, "MoreData");

// when the ExcelDataWriter is asynchronously disposed
// the buffered file is asynchronously written to the output.
```
