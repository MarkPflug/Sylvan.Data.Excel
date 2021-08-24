using Sylvan.Data.Csv;
using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using Xunit;

namespace Sylvan.Data.Excel
{
	public sealed class XlsxTests : ExcelTests
	{
		const string Format = "Data/Xlsx/{0}.xlsx";

		protected override string GetFile(string name)
		{
			return string.Format(Format, name);
		}
	}

	// the tests defined here will be run against both an .xls and .xlsx file
	// containing the same content. The expectation is
	public class ExcelTests
	{
		const string Format = "Data/Xls/{0}.xls";
		//const string Format = "data/xlsx/{0}.xlsx";

		protected virtual string GetFile([CallerMemberName] string name = "")
		{
			var file = string.Format(Format, name);
			Assert.True(File.Exists(file), "Test data file " + file + " does not exist");
			return file;
		}

		ExcelDataReaderOptions noHeaders =
			new ExcelDataReaderOptions
			{
				Schema = ExcelSchema.NoHeaders
			};

		public ExcelTests()
		{
			Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
		}

		[Fact]
		public void Big()
		{
			var file = GetFile();

			using var edr = ExcelDataReader.Create(file, noHeaders);
			for (int i = 0; i < 32; i++)
			{
				Assert.True(edr.Read());
				Assert.Equal(i + 1, edr.GetInt32(0));

				for (int j = 1; j < edr.FieldCount; j++)
				{
					Assert.Equal(j + 1, edr.GetInt32(j));
				}
			}
			// TODO: make this assertion pass
			// Assert.False(edr.Read());
		}

		[Fact]
		public void Numbers()
		{
			var file = GetFile();
			using var r = ExcelDataReader.Create(file, noHeaders);
			r.Read();
			Assert.Equal("3.3", r.GetString(0));
			Assert.Equal("1E+77", r.GetString(1));
			Assert.Equal("3.33", r.GetString(2));
			Assert.Equal("3.333", r.GetString(3));
			Assert.Equal("3.3333", r.GetString(4));
		}

		[Fact]
		public void DateTime()
		{
			var file = GetFile();

			var epoch = new DateTime(1900, 1, 1);
			using var edr = ExcelDataReader.Create(file);
			for (int i = 0; i < 22; i++)
			{
				Assert.True(edr.Read());
				var value = edr.GetDouble(0);
				var vs = value.ToString("G15");
				Assert.Equal(i / 10d, value);
				var v = edr.GetDouble(1);
				Assert.Equal(vs, v.ToString("G15"));
				v = edr.GetDouble(2);
				Assert.Equal(vs, v.ToString("G15"));
				v = edr.GetDouble(3);
				Assert.Equal(vs, v.ToString("G15"));

				if (i < 10)
				{
					var s = edr.GetString(1);
					Assert.Equal("", s);
					Assert.Throws<FormatException>(() => edr.GetDateTime(1));
					Assert.Throws<FormatException>(() => edr.GetDateTime(2));
					Assert.Throws<FormatException>(() => edr.GetDateTime(3));
				} 
				else
				{
					var dt = epoch.AddDays(value - 1);
					
					var dt1 = edr.GetDateTime(1);
					Assert.Equal(dt, dt1);
					var dt2 = edr.GetDateTime(2);
					Assert.Equal(dt, dt2);
					var dt3 = edr.GetDateTime(3);
					Assert.Equal(dt, dt3);
				}
			}
		}


		[Fact]
		public void FormatTest()
		{
			var file = GetFile("test");
			using var edr = ExcelDataReader.Create(file, noHeaders);
			var sw = new StringWriter();
			var csvW = CsvDataWriter.Create(sw);
			csvW.Write(edr);
			var str = sw.ToString();
		}

		

		[Fact]
		public void TestHeaders()
		{
			var file = GetFile("test");
			using var edr = ExcelDataReader.Create(file);
			var sw = new StringWriter();
			var csvW = CsvDataWriter.Create(sw);
			csvW.Write(edr);
			var str = sw.ToString();
		}

		[Fact]
		public void Gap()
		{
			var opts = new ExcelDataReaderOptions { GetErrorAsNull = true };
			var file = GetFile();
			using var edr = ExcelDataReader.Create(file, opts);
			var sw = new StringWriter();
			var csvW = CsvDataWriter.Create(sw);
			csvW.Write(edr);
			var str = sw.ToString();
		}

		[Fact]
		public void MultiSheet()
		{
			var opts = new ExcelDataReaderOptions { GetErrorAsNull = true };
			var file = GetFile();
			using var edr = ExcelDataReader.Create(file, opts);
			Assert.Equal(2, edr.WorksheetCount);
			Assert.Equal("Primary", edr.WorksheetName);
			Assert.True(edr.NextResult());
			Assert.Equal("Secondary", edr.WorksheetName);
			Assert.False(edr.NextResult());
		}

		[Fact]
		public void Func()
		{
			var opts =
				new ExcelDataReaderOptions
				{
					GetErrorAsNull = true,
					Schema = ExcelSchema.NoHeaders
				};

			var file = GetFile();

			using var edr = ExcelDataReader.Create(file, opts);
			Assert.Equal(3, edr.FieldCount);

			Assert.True(edr.Read()); // 1
			Assert.Equal(ExcelDataType.Numeric, edr.GetExcelDataType(0));
			Assert.Equal(0, edr.GetDouble(0));
			Assert.Equal(ExcelDataType.Boolean, edr.GetExcelDataType(1));
			Assert.True(edr.GetBoolean(1));
			Assert.Equal(ExcelDataType.Error, edr.GetExcelDataType(2));
			Assert.Equal(ExcelErrorCode.DivideByZero, edr.GetFormulaError(2));
			var ex = Assert.Throws<ExcelFormulaException>(() => edr.GetDouble(2));
			Assert.Equal(ExcelErrorCode.DivideByZero, ex.ErrorCode);

			Assert.True(edr.Read()); // 2
			Assert.Equal(ExcelDataType.Numeric, edr.GetExcelDataType(0));
			Assert.Equal(1, edr.GetDouble(0));
			Assert.Equal(ExcelDataType.Boolean, edr.GetExcelDataType(1));
			Assert.False(edr.GetBoolean(1));
			Assert.Equal(ExcelDataType.Numeric, edr.GetExcelDataType(2));
			Assert.Equal(2, edr.GetDouble(2));

			Assert.True(edr.Read()); // 3
			Assert.Equal(ExcelDataType.Numeric, edr.GetExcelDataType(0));
			Assert.Equal(2, edr.GetDouble(0));
			Assert.Equal(ExcelDataType.Boolean, edr.GetExcelDataType(1));
			Assert.True(edr.GetBoolean(1));
			Assert.Equal(ExcelDataType.Numeric, edr.GetExcelDataType(2));
			Assert.Equal(1.5, edr.GetDouble(2));

			Assert.True(edr.Read()); // 4
			Assert.Equal(ExcelDataType.Numeric, edr.GetExcelDataType(0));
			Assert.Equal(3, edr.GetDouble(0));
			Assert.Equal(ExcelDataType.Boolean, edr.GetExcelDataType(1));
			Assert.False(edr.GetBoolean(1));
			Assert.Equal(ExcelDataType.Numeric, edr.GetExcelDataType(2));
			Assert.Equal(2, edr.GetDouble(2));

			Assert.True(edr.Read()); // 5
			Assert.Equal(ExcelDataType.Numeric, edr.GetExcelDataType(0));
			Assert.Equal(6, edr.GetDouble(0));
			Assert.Equal(ExcelDataType.Boolean, edr.GetExcelDataType(1));
			Assert.True(edr.GetBoolean(1));
			Assert.Equal(ExcelDataType.Error, edr.GetExcelDataType(2));
			ex = Assert.Throws<ExcelFormulaException>(() => edr.GetDouble(2));
			Assert.Equal(ExcelErrorCode.Value, ex.ErrorCode);

			Assert.True(edr.Read()); // 6
			Assert.Equal(ExcelDataType.String, edr.GetExcelDataType(0));
			Assert.Equal("a", edr.GetString(0));
			Assert.Equal(ExcelDataType.String, edr.GetExcelDataType(1));
			Assert.Equal("b", edr.GetString(1));
			Assert.Equal(ExcelDataType.String, edr.GetExcelDataType(2));
			Assert.Equal("ab", edr.GetString(2));
		}
	}
}
