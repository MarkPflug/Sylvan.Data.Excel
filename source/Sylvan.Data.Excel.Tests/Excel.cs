using Sylvan.Data.Csv;
using System.IO;
using System.Text;
using Xunit;

namespace Sylvan.Data.Excel
{
	public class Excel
	{
		const string FileName = @"\data\excel\Excel Pkdx V5.14.2.xlsx";
		//const string FileName = @"\data\excel\itcont.xlsx";
		//const string XlsFileName = @"\data\excel\Excel Pkdx V5.14.xls";

		public Excel()
		{
			Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
		}

		[Fact]
		public void Test1()
		{
			var reader = ExcelDataReader.Create(FileName);
			var sw = new StringWriter();
			var csv = CsvDataWriter.Create(sw);
			csv.Write(reader);
			var data = sw.ToString();
		}

		[Fact]
		public void Test2()
		{
			var reader = ExcelDataReader.Create(FileName);
			var sw = new StringWriter();
			var csv = CsvDataWriter.Create(sw);
			csv.Write(reader);
			//do
			//{
			//	while (reader.Read())
			//	{
			//		for (int i = 0; i < reader.FieldCount; i++)
			//		{
			//			reader.GetString(i);
			//		}
			//	}
			//} while (reader.NextResult());
			var str = sw.ToString();
		}

		[Fact]
		public void TestCsv()
		{
			using var edr = ExcelDataReader.Create("/data/excel/65K_Records_Data.xls");
			do
			{
				var sheetName = edr.WorksheetName;
				using var cdw = CsvDataWriter.Create("data-" + sheetName + ".csv");
				cdw.Write(edr);
			} while (edr.NextResult());
		}

		[Fact]
		public void TestBig()
		{
			using var edr = ExcelDataReader.Create("data/xls/big.xls");
			var sw = new StringWriter();
			var csv = CsvDataWriter.Create(sw, new CsvDataWriterOptions { BufferSize = 0x100000 });
			do
			{
				csv.Write(edr);
			} while (edr.NextResult());
			var str = sw.ToString();
		}

		[Fact]
		public void FormatTest()
		{
			using var edr = ExcelDataReader.Create("data/xls/test.xls", new ExcelDataReaderOptions { Schema = ExcelSchema.NoHeaders });
			var sw = new StringWriter();
			var csvW = CsvDataWriter.Create(sw);
			csvW.Write(edr);
			var str = sw.ToString();
		}

		[Fact]
		public void DateTimeTest()
		{
			using var edr = ExcelDataReader.Create("data/xls/datetime.xls");
			var sw = new StringWriter();
			var csvW = CsvDataWriter.Create(sw);
			csvW.Write(edr);
			var str = sw.ToString();
		}

		[Fact]
		public void TestHeaders()
		{
			using var edr = ExcelDataReader.Create("data/xls/test.xls");
			var sw = new StringWriter();
			var csvW = CsvDataWriter.Create(sw);
			csvW.Write(edr);
			var str = sw.ToString();
		}

		[Fact]
		public void TestFuncs()
		{
			var opts = new ExcelDataReaderOptions { GetErrorAsNull = true };
			using var edr = ExcelDataReader.Create("data/xls/func.xls", opts);
			var sw = new StringWriter();
			var csvW = CsvDataWriter.Create(sw);
			csvW.Write(edr);
			var str = sw.ToString();
		}

		[Fact]
		public void TestGap()
		{
			var opts = new ExcelDataReaderOptions { GetErrorAsNull = true };
			using var edr = ExcelDataReader.Create("data/xls/gap.xls", opts);
			var sw = new StringWriter();
			var csvW = CsvDataWriter.Create(sw);
			csvW.Write(edr);
			var str = sw.ToString();
		}

		[Fact]
		public void MultiSheet()
		{
			var opts = new ExcelDataReaderOptions { GetErrorAsNull = true };
			using var edr = ExcelDataReader.Create("data/xls/multiSheet.xls", opts);
			var sw = new StringWriter();
			var csvW = CsvDataWriter.Create(sw);
			csvW.Write(edr);
			edr.NextResult();
			csvW.Write(edr);
			var str = sw.ToString();
		}

		[Fact]
		public void TestFuncsXlsx()
		{
			var opts = 
				new ExcelDataReaderOptions { 
				GetErrorAsNull = true, 
				Schema = ExcelSchema.NoHeaders 
			};
			using var edr = ExcelDataReader.Create("data/xlsx/func.xlsx", opts);
			Assert.Equal(3, edr.FieldCount);
			
			Assert.True(edr.Read());
			Assert.Equal(ExcelDataType.Numeric, edr.GetExcelDataType(0));
			Assert.Equal(0, edr.GetDouble(0));
			Assert.Equal(ExcelDataType.Boolean, edr.GetExcelDataType(1));
			Assert.True(edr.GetBoolean(1));
			Assert.Equal(ExcelDataType.Error, edr.GetExcelDataType(2));
			Assert.Equal(ExcelErrorCode.DivideByZero, edr.GetFormulaError(2));
			var ex = Assert.Throws<ExcelFormulaException>(() => edr.GetDouble(2));
			Assert.Equal(ExcelErrorCode.DivideByZero, ex.ErrorCode);

			Assert.True(edr.Read());
			Assert.Equal(ExcelDataType.Numeric, edr.GetExcelDataType(0));
			Assert.Equal(1, edr.GetDouble(0));
			Assert.Equal(ExcelDataType.Boolean, edr.GetExcelDataType(1));
			Assert.False(edr.GetBoolean(1));
			Assert.Equal(ExcelDataType.Numeric, edr.GetExcelDataType(2));
			Assert.Equal(2, edr.GetDouble(2));

			Assert.True(edr.Read());
			Assert.Equal(ExcelDataType.Numeric, edr.GetExcelDataType(0));
			Assert.Equal(2, edr.GetDouble(0));
			Assert.Equal(ExcelDataType.Boolean, edr.GetExcelDataType(1));
			Assert.True(edr.GetBoolean(1));
			Assert.Equal(ExcelDataType.Numeric, edr.GetExcelDataType(2));
			Assert.Equal(1.5, edr.GetDouble(2));

			Assert.True(edr.Read());
			Assert.Equal(ExcelDataType.Numeric, edr.GetExcelDataType(0));
			Assert.Equal(3, edr.GetDouble(0));
			Assert.Equal(ExcelDataType.Boolean, edr.GetExcelDataType(1));
			Assert.False(edr.GetBoolean(1));
			Assert.Equal(ExcelDataType.Numeric, edr.GetExcelDataType(2));
			Assert.Equal(2, edr.GetDouble(2));

			Assert.True(edr.Read());
			Assert.Equal(ExcelDataType.Numeric, edr.GetExcelDataType(0));
			Assert.Equal(6, edr.GetDouble(0));
			Assert.Equal(ExcelDataType.Boolean, edr.GetExcelDataType(1));
			Assert.True(edr.GetBoolean(1));
			Assert.Equal(ExcelDataType.Error, edr.GetExcelDataType(2));
			ex = Assert.Throws<ExcelFormulaException>(() => edr.GetDouble(2));
			Assert.Equal(ExcelErrorCode.Value, ex.ErrorCode);

			Assert.True(edr.Read());
			Assert.Equal(ExcelDataType.String, edr.GetExcelDataType(0));
			Assert.Equal("a", edr.GetString(0));
			Assert.Equal(ExcelDataType.String, edr.GetExcelDataType(1));
			Assert.Equal("b", edr.GetString(1));
			Assert.Equal(ExcelDataType.String, edr.GetExcelDataType(2));
			Assert.Equal("ab", edr.GetString(2));

			Assert.True(edr.Read());
			Assert.Equal(ExcelDataType.Null, edr.GetExcelDataType(0));
			Assert.Equal(ExcelDataType.Null, edr.GetExcelDataType(1));
			Assert.Equal(ExcelDataType.Error, edr.GetExcelDataType(2));
			Assert.Equal(ExcelErrorCode.Reference, edr.GetFormulaError(2));

			Assert.True(edr.Read());
			Assert.Equal(ExcelDataType.Null, edr.GetExcelDataType(0));
			Assert.Equal(ExcelDataType.Null, edr.GetExcelDataType(1));
			Assert.Equal(ExcelDataType.Error, edr.GetExcelDataType(2));
			Assert.Equal(ExcelErrorCode.Name, edr.GetFormulaError(2));

			Assert.True(edr.Read());
			Assert.Equal(ExcelDataType.Null, edr.GetExcelDataType(0));
			Assert.Equal(ExcelDataType.Null, edr.GetExcelDataType(1));
			Assert.Equal(ExcelDataType.Error, edr.GetExcelDataType(2));
			Assert.Equal(ExcelErrorCode.Null, edr.GetFormulaError(2));
		}
	}
}
