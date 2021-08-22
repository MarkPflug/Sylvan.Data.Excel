using Sylvan.Data.Csv;
using System.IO;
using System.Text;
using Xunit;

namespace Sylvan.Data.Excel
{
	public class ExcelDataReaderTests
	{
		static ExcelDataReaderTests()
		{
			Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
		}

		[Fact]
		public void TestBig()
		{
			//var file = @"\data\excel\65K_Records_Data.xls";
			var file = "/data/excel/vehicles_xls.xls";

			using var r = ExcelDataReader.Create(file);
			while (r.Read()) ;
			//var str = r.GetString(21);
			//using var w = CsvDataWriter.Create(Path.GetFileName(file) + ".csv", new CsvDataWriterOptions { BufferSize = 0x1000000 });
			//w.Write(r);
		}

		[Fact]
		public void Test1()
		{
			using var r = ExcelDataReader.Create("C:/users/mark/downloads/pkdx.xls");

			while (r.Read()) { }
			Assert.True(r.NextResult());
			while (r.Read()) { }
			Assert.True(r.NextResult());
			while (r.Read()) { }
		}

		[Fact]
		public void Test2()
		{
			using var r = ExcelDataReader.Create("C:/users/mark/downloads/pkdx.xls");

			Assert.True(r.NextResult());
			while (r.Read()) { }
			Assert.True(r.NextResult());
			while (r.Read()) { }
		}

		[Fact]
		public void Test3()
		{
			//using var r = ExcelDataReader.Create("C:/users/mark/downloads/pkdx.xls");
			using var r = ExcelDataReader.Create("C:/users/mark/downloads/us-mr2010-01.xls");
			var sw = new StringWriter();
			using (var csv = CsvDataWriter.Create(sw))
			{
				do
				{
					csv.Write(r);

					sw.WriteLine();
				} while (r.NextResult());
				sw.WriteLine();
			}
			var str = sw.ToString();			
		}

		[Fact]
		public void Numbers()
		{
			using var r = ExcelDataReader.Create("Data/Xls/Numbers.xls", new ExcelDataReaderOptions { Schema = ExcelSchema.NoHeaders });
			r.Read();
			Assert.Equal("3.3", r.GetString(0));
			Assert.Equal("1E+77", r.GetString(1));
			Assert.Equal("3.33", r.GetString(2));
			Assert.Equal("3.333", r.GetString(3));
			Assert.Equal("3.3333", r.GetString(4));
		}
	}
}
