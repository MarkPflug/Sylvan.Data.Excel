using System;
using System.Data;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using Xunit;

namespace Sylvan.Data.Excel
{
	public sealed class XlsxTests : ExcelTests
	{
		const string FileFormat = "Data/{0}.xlsx";

		protected override string GetFile(string name)
		{
			return string.Format(FileFormat, name);
		}
	}

	public sealed class XlsbTests : ExcelTests
	{
		const string FileFormat = "Data/{0}.xlsb";

		protected override string GetFile(string name)
		{
			return string.Format(FileFormat, name);
		}
	}

	// the tests defined here will be run against .xls, .xlsx, and .xlsb file
	// containing the same content. The expectation is the behavior of the two
	// implementations is the same, so the same test code can validate the 
	// behavior of the three formats.
	public class ExcelTests
	{
		const string FileFormat = "Data/{0}.xls";

		protected virtual string GetFile([CallerMemberName] string name = "")
		{
			var file = string.Format(FileFormat, name);
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
#if NET6_0_OR_GREATER
			Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
#endif
		}

		[Fact]
		public void Big()
		{
			var file = GetFile();

			using var edr = ExcelDataReader.Create(file, noHeaders);
			for (int i = 0; i < 32; i++)
			{
				Assert.True(edr.Read(), "Row " + i);
				Assert.Equal(i + 1, edr.GetInt32(0));

				for (int j = 1; j < edr.FieldCount; j++)
				{
					Assert.Equal(j + 1, edr.GetInt32(j));
				}
			}
			Assert.False(edr.Read());
		}

		static readonly string[] ColumnHeaders = new[]
		{
			"Id", "Name", "Date", "Amount", "Code", "Flagged", "Lat", "Lon"
		};

		[Fact]
		public void Headers()
		{
			var file = GetFile("Schema");
			using var r = ExcelDataReader.Create(file);

			for (int i = 0; i < r.FieldCount; i++)
			{
				Assert.Equal(ColumnHeaders[i], r.GetName(i));
			}
		}

		[Fact]
		public void HeadersWithSchema()
		{
			var file = GetFile("Schema");
			var schema = new ExcelSchema(true, GetSchema());
			using var r = ExcelDataReader.Create(file, new ExcelDataReaderOptions { Schema = schema });

			for (int i = 0; i < r.FieldCount; i++)
			{
				Assert.Equal(ColumnHeaders[i], r.GetName(i));
			}
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
			Assert.False(r.Read());
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
					Assert.Equal(System.DateTime.MinValue.AddMinutes(i * 144), edr.GetDateTime(3));
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
			Assert.False(edr.Read());
		}

		[Fact]
		public void Format()
		{
			var file = GetFile();
			using var edr = ExcelDataReader.Create(file, noHeaders);
			int row = 0;
			ExcelFormat fmt;
			for (int i = 0; i < 20; i++)
			{
				Assert.True(edr.Read());
				Assert.Equal(i + 1, edr.GetInt32(0));
				fmt = edr.GetFormat(1);
				if (!edr.IsDBNull(1))
					Assert.Equal(FormatKind.Number, fmt.Kind);
				fmt = edr.GetFormat(2);
				if (!edr.IsDBNull(2))
					Assert.Equal(FormatKind.Date, fmt.Kind);
				fmt = edr.GetFormat(3);
				if (!edr.IsDBNull(3))
					Assert.Equal(FormatKind.Time, fmt.Kind);
				row++;
			}
			Assert.False(edr.Read());
		}

		[Fact]
		public void Gap()
		{
			var file = GetFile();
			using var edr = ExcelDataReader.Create(file, noHeaders);
			for (int i = 0; i < 41; i++)
			{
				edr.Read();
				var str = edr.GetString(0);
				if (i % 10 == 0)
				{
					Assert.Equal("" + ((char)('a' + i / 10)), str);
				}
				else
				{
					Assert.Equal("", str);
				}
			}
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
			Assert.Null(edr.WorksheetName);
		}

		Schema GetSchema(string name = "Schema")
		{
			var schemaText = File.ReadAllText("Data/" + name + ".txt");
			var schema = Data.Schema.Parse(schemaText);
			return schema;
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
			Assert.False(edr.Read());
		}

		[Fact]
		public void GetValueTest()
		{
			var file = GetFile("Schema");

			using var edr = ExcelDataReader.Create(file);
			while (edr.Read())
			{
				for (int i = 0; i < edr.FieldCount; i++)
				{
					var value = edr.GetValue(i);
					Assert.True(value is string);
				}
			}
		}

		[Fact]
		public void GetValueWithSchemaTest()
		{
			var schema = GetSchema();
			var opts = new ExcelDataReaderOptions { Schema = new ExcelSchema(true, schema) };
			var file = GetFile("Schema");
			using var edr = ExcelDataReader.Create(file, opts);
			var cols = schema.GetColumnSchema();
			while (edr.Read())
			{
				for (int i = 0; i < edr.FieldCount; i++)
				{
					var value = edr.GetValue(i);

					Assert.IsType(cols[i].DataType, value);
				}
			}
		}

		[Fact]
		public void Schema()
		{
			var schema = GetSchema();
			var opts = new ExcelDataReaderOptions { Schema = new ExcelSchema(true, schema) };
			var file = GetFile();
			using var edr = ExcelDataReader.Create(file, opts);

			Assert.Equal(typeof(int), edr.GetFieldType(0));
			Assert.Equal(typeof(string), edr.GetFieldType(1));
			Assert.Equal(typeof(DateTime), edr.GetFieldType(2));
			Assert.Equal(typeof(decimal), edr.GetFieldType(3));
			Assert.Equal(typeof(string), edr.GetFieldType(4));
			Assert.Equal(typeof(bool), edr.GetFieldType(5));
			Assert.Equal(typeof(double), edr.GetFieldType(6));
			Assert.Equal(typeof(double), edr.GetFieldType(7));

			var colSchema = edr.GetColumnSchema();
			for (int i = 0; i < colSchema.Count; i++)
			{
				Assert.Equal(colSchema[i].DataType, edr.GetFieldType(i));
			}

			var names = new[] { "James", "Janet", "Frank", "Laura" };
			var dates = new[] {
				new DateTime(2020, 1, 1),
				new DateTime(2022, 1, 1),
				new DateTime(2021, 1, 1),
				new DateTime(2019, 1, 1),
			};

			for (int i = 0; i < 4; i++)
			{
				Assert.True(edr.Read());
				Assert.Equal(i + 1, edr.GetInt32(0));
				Assert.Equal(i + 1, edr.GetInt16(0));
				Assert.Equal(i + 1, edr.GetValue(0));
				Assert.Equal("" + (i + 1), edr.GetString(0));
				Assert.Equal(names[i], edr.GetString(1));

				Assert.Equal(dates[i], edr.GetDateTime(2));

				Assert.Equal(i >= 2, edr.GetBoolean(5));
				Assert.Equal((i >= 2).ToString(), edr.GetString(5));

				var a = edr.GetDouble(3);
				var b = edr.GetDecimal(3);
				Assert.Equal(a, (double)b);

				a = edr.GetDouble(6);
				b = edr.GetDecimal(6);
				Assert.Equal(a, (double)b);

				a = edr.GetDouble(7);
				b = edr.GetDecimal(7);
				Assert.Equal(a, (double)b);

			}
			Assert.False(edr.Read());
		}

		[Fact]
		public void GetSchemaTable()
		{
			var schema = GetSchema();
			var opts = new ExcelDataReaderOptions { Schema = new ExcelSchema(true, schema) };
			var file = GetFile("Schema");
			using var edr = ExcelDataReader.Create(file, opts);
			var st = edr.GetSchemaTable();
			Assert.Equal(8, st.Rows.Count);
		}

		[Fact]
		public void DataTable()
		{
			var schema = GetSchema();
			var opts = new ExcelDataReaderOptions { Schema = new ExcelSchema(true, schema) };
			var file = GetFile("Schema");
			using var edr = ExcelDataReader.Create(file, opts);
			var dt = new DataTable();
			dt.Load(edr);
			Assert.Equal(4, dt.Rows.Count);
		}

		[Fact]
		public void Jagged()
		{
			var file = GetFile();
			using var edr = ExcelDataReader.Create(file);

			Assert.Equal(3, edr.FieldCount);
			Assert.True(edr.Read());
			Assert.Equal(3, edr.RowFieldCount);
			Assert.True(edr.Read());
			Assert.Equal(2, edr.RowFieldCount);
			Assert.True(edr.Read());
			Assert.Equal(1, edr.RowFieldCount);
			Assert.True(edr.Read());
			Assert.Equal(3, edr.RowFieldCount);
			Assert.True(edr.Read());
			Assert.Equal(4, edr.RowFieldCount);
			Assert.True(edr.Read());
			Assert.Equal(5, edr.RowFieldCount);
			Assert.False(edr.Read());
		}
	}
}
