using System.Data.Common;
using System.IO;
using System.Threading.Tasks;
using Xunit;

namespace Sylvan.Data.Excel;

public partial class XlsxWriterAsyncTests
{
	public static DbDataReader GetTestData()
	{
		var records = new[]
		{
			new
			{
				Id = 1,
				Name = "Test",
				Value = 12.25,
			},
			new
			{
				Id = 2,
				Name = "Qwerty",
				Value = 13.99,
			},
			new
			{
				Id = 3,
				Name = "Dvorak",
				Value = 17.5,
			}
		};
		return records.AsDataReader();
	}

	[Fact]
	public async Task TestAsync()
	{
		var data = GetTestData();
		var ms = new MemoryStream();
		{
			await using var w = await ExcelDataWriter.CreateAsync(ms, ExcelWorkbookType.ExcelXml);
			await w.WriteAsync(data);
		}
		ms.Seek(0, SeekOrigin.Begin);
		Assert.NotEqual(0, ms.Length);

		using var r = await ExcelDataReader.CreateAsync(ms, ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions { Schema = ExcelSchema.Dynamic });

		var b = GetTestData();

		while (await r.ReadAsync())
		{
			Assert.True(b.Read());
			for (int i = 0; i < r.FieldCount; i++)
			{
				Assert.Equal(b.GetValue(i), r.GetValue(i));
			}
		}
	}
}
