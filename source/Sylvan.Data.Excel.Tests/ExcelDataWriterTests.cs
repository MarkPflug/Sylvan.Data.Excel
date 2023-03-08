using Sylvan.Data.Csv;
using System;
using System.Collections;
using System.Data.Common;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.CompilerServices;
using Xunit;

namespace Sylvan.Data.Excel;

public class XlsxDataWriterTests : ExcelDataWriterTests
{
	const string FileFormat = "{0}.xlsx";

	public override ExcelWorkbookType WorkbookType => ExcelWorkbookType.Excel;

	protected override string GetFile(string name)
	{
		return string.Format(FileFormat, name);
	}
}

public class XlsbDataWriterTests : ExcelDataWriterTests
{
	const string FileFormat = "{0}.xlsb";

	public override ExcelWorkbookType WorkbookType => ExcelWorkbookType.ExcelBinary;

	protected override string GetFile(string name)
	{
		return string.Format(FileFormat, name);
	}
}

public abstract class ExcelDataWriterTests
{
	protected abstract string GetFile([CallerMemberName] string name = null);

	public abstract ExcelWorkbookType WorkbookType { get; }
	public object Enumable { get; private set; }

	static void Unpack(string file, [CallerMemberName] string folder = null)
	{
		// useful for debugging.
		try
		{
			Directory.Delete(folder, true);
		}
		catch { }
		ZipFile.ExtractToDirectory(file, Path.GetDirectoryName(file) + folder);
	}

	static void Open(string file)
	{
		var psi = new ProcessStartInfo(file)
		{
			UseShellExecute = true,
		};
		Process.Start(psi);
	}

	[Fact]
	public void Simple()
	{
		// tests the most common types.
		Random r = new Random();
		var data =
			Enumerable.Range(1, 4096)
			.Select(
				i => new
				{
					Id = i, //int32
					Name = "Name" + i, //string
					ValueInt = r.Next(), // another, bigger int
					ValueDouble = Math.PI * i, // double
					Decimal = 1.25m * i,
					Date = DateTime.Today.AddHours(i),
				}
			);

		var f = GetFile();
		var reader = data.AsDataReader();
		using (var w = ExcelDataWriter.Create(f))
		{
			w.Write(reader);
		}
		Open(f);
	}

	[Fact]
	public void Decimal()
	{
		// tests the most common types.
		Random r = new Random();
		var data =
			new[]
			{
				decimal.MinValue,
				decimal.MaxValue,
				(decimal)int.MinValue,
				(decimal)int.MaxValue,
				-(decimal)(1 << 24),
				(decimal)(1 << 24),
				-1m,
				1m,
				6266593.83m,
			}.Select(v => new { Decimal = v });

		var f = GetFile();
		var reader = data.AsDataReader();
		using (var w = ExcelDataWriter.Create(f))
		{
			w.Write(reader);
		}
		Open(f);
	}

	[Fact]
	public void Ints()
	{
		// tests the most common types.
		Random r = new Random();
		var data =
			new[]
			{
				int.MinValue,
				(int) short.MinValue,
				-1,
				0,
				1,
				(int) short.MaxValue,
				int.MaxValue,
			}.Select(v => new { Value = v });

		var f = GetFile();
		var reader = data.AsDataReader();
		using (var w = ExcelDataWriter.Create(f))
		{
			w.Write(reader);
		}
		Open(f);
	}

	[Fact]
	public void CommonTypes()
	{
		// tests the most common types.
		Random r = new Random();
		var data =
			Enumerable.Range(1, 100)
			.Select(
				i => new
				{
					Id = i, //int32
					Name = "Name" + i, //string
					ValueInt = r.Next(), // another, bigger int
					ValueDouble = r.NextDouble() * 100d, // double
					Amount = (decimal)r.NextDouble(), // decimal
					DateTime = new DateTime(2020, 1, 1).AddHours(i), // datetime
				}
			);

		var f = GetFile();
		var reader = data.AsDataReader();
		using (var w = ExcelDataWriter.Create(f))
		{
			w.Write(reader);
		}
		//Open(f);
	}

	[Fact]
	public void WorkSheetNameSize()
	{
		var data = Enumerable.Range(1, 100).Select(i => new { Id = i, Name = "Name" + i });

		var f = GetFile();

		using (var w = ExcelDataWriter.Create(f))
		{
			Assert.Throws<ArgumentException>(() => w.Write(data.AsDataReader(), new string('a', 32)));
		}
	}

	[Fact]
	public void MultiSheet()
	{
		Random r = new Random();
		var data =
			Enumerable.Range(1, 100)
			.Select(
				i => new
				{
					Id = i, //int32
					Name = "Name" + i, //string
					ValueInt = r.Next(), // another, bigger int
					ValueDouble = r.NextDouble() * 100d, // double
					Amount = (decimal)r.NextDouble(), // decimal
					DateTime = new DateTime(2020, 1, 1).AddHours(i), // datetime
				}
			);

		var f = GetFile();

		using (var w = ExcelDataWriter.Create(f))
		{

			var reader = data.AsDataReader();
			w.Write(reader);

			reader = data.AsDataReader();
			w.Write(reader);
		}
		Unpack(f);
	}

	[Fact]
	public void BigString()
	{
		// this is the largest string that can be written.
		// larger, and Excel will complain, and truncate it.
		var bigStr = new string('a', short.MaxValue);
		var data =
			Enumerable.Range(1, 10)
			.Select(
				i => new
				{
					Id = i, //int32
					BigString = bigStr
				}
			);

		var f = GetFile();
		var reader = data.AsDataReader();
		using (var w = ExcelDataWriter.Create(f))
		{
			w.Write(reader);
		}
		//Open(f);
	}

	[Fact]
	public void NullCharString()
	{
		var str = "a\0b";
		var data =
			Enumerable.Range(1, 10)
			.Select(
				i => new
				{
					Id = i, //int32
					String = str
				}
			);

		var f = GetFile();
		var reader = data.AsDataReader();
		using (var w = ExcelDataWriter.Create(f))
		{
			w.Write(reader);
		}
		//Open(f);
	}

	[Fact]
	public void Boolean()
	{
		var data =
			Enumerable.Range(1, 100)
			.Select(
				i => new
				{
					Id = i, //int32
					Boolean = (i & 1) != 0
				}
			);

		var f = GetFile();
		var reader = data.AsDataReader();
		using (var w = ExcelDataWriter.Create(f))
		{
			w.Write(reader);
		}
		//Open(f);
	}

	[Fact]
	public void JaggedData()
	{
		// tests writing jagged data to excel.
		var data = "a,b,c\n1,2,3\n1,2,3,4\n,1,2,3,4,5\n";
		var r = new StringReader(data);
		var csv = CsvDataReader.Create(r);

		var dr = csv.AsVariableField(c => c.RowFieldCount);
		var f = GetFile();
		using (var edw = ExcelDataWriter.Create(f))
		{
			edw.Write(dr);
		}
		//Open(f);
	}

	[Fact]
	public void WhiteSpace()
	{
		var data = " a , b,c \n 1 , 2,3 \n";
		var r = new StringReader(data);
		var csv = CsvDataReader.Create(r);

		var f = GetFile();
		using (var edw = ExcelDataWriter.Create(f))
		{
			edw.Write(csv);
		}

		using var edr = ExcelDataReader.Create(f);
		Assert.Equal(" a ", edr.GetName(0));
		Assert.Equal(" b", edr.GetName(1));
		Assert.Equal("c ", edr.GetName(2));
		Assert.True(edr.Read());
		Assert.Equal(" 1 ", edr.GetString(0));
		Assert.Equal(" 2", edr.GetString(1));
		Assert.Equal("3 ", edr.GetString(2));
		Assert.False(edr.Read());

	}

	[Fact]
	public void Binary()
	{
		var data = new[]
		{
			new {
				Name = "A",
				Data = new byte[] {1,2,3,4 },
			},
			new {
				Name = "B",
				Data = new byte[] {1,2,3,4 },
			}
		};

		var dr = data.AsDataReader();
		var f = GetFile();
		using (var edw = ExcelDataWriter.Create(f))
		{
			edw.Write(dr);
		}
		//Open(f);
	}

	[Fact]
	public void CharArray()
	{
		var dr = new CharTestDataReader();
		var f = GetFile();
		using (var edw = ExcelDataWriter.Create(f))
		{
			edw.Write(dr);
		}

		// read back the created file and assert everything is as we expected
		using (var edr = ExcelDataReader.Create(f))
		{
			Assert.Equal("Name", edr.GetName(0));
			Assert.Equal("Data", edr.GetName(1));
			Assert.True(edr.Read());
			Assert.Equal("a", edr.GetString(0));
			Assert.Equal("alphabet", edr.GetString(1));
			Assert.True(edr.Read());
			Assert.Equal("b", edr.GetString(0));
			Assert.Equal(new string('z', short.MaxValue), edr.GetString(1));
			Assert.False(edr.Read());
		}
	}

	[Fact]
	public void Char()
	{
		var data = new[]
		{
			new {
				Name = "Alpha",
				Data = 'A',
			},
			new {
				Name = "Beta",
				Data = 'B',
			}
		};

		var dr = data.AsDataReader();
		var f = GetFile();
		using (var edw = ExcelDataWriter.Create(f))
		{
			edw.Write(dr);
		}
		//Open(f);
	}

	[Fact]
	public void Byte()
	{
		var data = new[]
		{
			new {
				Name = "Alpha",
				Data = (byte)'A',
			},
			new {
				Name = "Beta",
				Data = (byte)'B',
			}
		};

		var dr = data.AsDataReader();
		var f = GetFile();
		using (var edw = ExcelDataWriter.Create(f))
		{
			edw.Write(dr);
		}
		//Open(f);
	}

	[Fact]
	public void GuidData()
	{
		var data =
			Enumerable.Range(0, 16)
			.Select(
				i =>
				{
					var b = new byte[16];
					Array.Fill(b, (byte)(i | i << 4));
					return new
					{
						Name = "Id " + i,
						Guid = new Guid(b),
						Value = i,
					};
				}
			);

		var dr = data.AsDataReader();
		var f = GetFile();
		using (var edw = ExcelDataWriter.Create(f))
		{
			edw.Write(dr);
		}
		//Open(f);
	}

	[Fact]
	public void TimeSpanData()
	{
		var data =
			Enumerable.Range(0, 100)
			.Select(
				i =>
				{
					return new
					{
						TimeSpan = TimeSpan.FromSeconds(Math.PI * i),
						Value = i,
					};
				}
			);

		var dr = data.AsDataReader();
		var f = GetFile();
		using (var edw = ExcelDataWriter.Create(f))
		{
			edw.Write(dr);
		}
		//Open(f);
	}

#if NET6_0_OR_GREATER

	[Fact]
	public void DateOnly()
	{
		var data =
			Enumerable.Range(1, 100)
			.Select(
				i => new
				{
					Id = i,
					DateOnly = new DateOnly(2020, 1, 1).AddDays(i),
				}
			);

		var f = GetFile();
		var reader = data.AsDataReader();
		using (var w = ExcelDataWriter.Create(f))
		{
			w.Write(reader);
		}
		//Open(f);
	}

	[Fact]
	public void TimeOnly()
	{
		var data =
			Enumerable.Range(1, 100)
			.Select(
				i => new
				{
					Id = i,
					TimeOnly = new TimeOnly(1, 0).AddMinutes(i * 7),
				}
			);

		var f = GetFile();
		var reader = data.AsDataReader();
		using (var w = ExcelDataWriter.Create(f))
		{
			w.Write(reader);
		}
		//Open(f);
	}

#endif
}

// TODO: this is used to test writing char[] values.
// need to fix System.Data.ObjectDataReader to replace this.
class CharTestDataReader : DbDataReader
{
	const int RowCount = 2;
	const int ColCount = 2;

	int row = -1;

	string[] names;
	string[] col0;
	char[][] col1;

	public CharTestDataReader()
	{
		this.names = new[] { "Name", "Data" };
		this.col0 =
			new[] {
				"a",
				"b"
			};
		this.col1 =
			new[] {
				"alphabet".ToCharArray(),
				new string('z', short.MaxValue).ToCharArray()
			};
	}

	public override int FieldCount => ColCount;

	public override Type GetFieldType(int ordinal)
	{
		switch (ordinal)
		{
			case 0: return typeof(string);
			case 1: return typeof(char[]);
		}
		throw new ArgumentOutOfRangeException();
	}

	public override int GetInt32(int ordinal) => row + ordinal;

	public override bool IsDBNull(int ordinal) => false;

	public override string GetName(int ordinal) => names[ordinal];

	public override string GetString(int ordinal)
	{
		return col0[row];
	}

	public override long GetChars(int ordinal, long dataOffset, char[] buffer, int bufferOffset, int length)
	{
		var dat = col1[row];
		var count = Math.Min(length, dat.Length - (int)dataOffset);

		dat.AsSpan((int)dataOffset, count).CopyTo(buffer.AsSpan(bufferOffset));
		return count;
	}

	public override bool Read()
	{
		row++;
		return row < RowCount;
	}

	public override object GetValue(int ordinal)
	{
		return
			ordinal % 2 == 0
			? (object)GetString(ordinal)
			: (object)GetInt32(ordinal);
	}

	#region NotImplemented

	public override object this[int ordinal] => throw new NotImplementedException();

	public override object this[string name] => throw new NotImplementedException();

	public override int Depth => throw new NotImplementedException();

	public override bool HasRows => throw new NotImplementedException();

	public override bool IsClosed => throw new NotImplementedException();

	public override int RecordsAffected => throw new NotImplementedException();

	public override bool GetBoolean(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override byte GetByte(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override long GetBytes(int ordinal, long dataOffset, byte[] buffer, int bufferOffset, int length)
	{
		throw new NotImplementedException();
	}

	public override char GetChar(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override string GetDataTypeName(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override DateTime GetDateTime(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override decimal GetDecimal(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override double GetDouble(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override IEnumerator GetEnumerator()
	{
		throw new NotImplementedException();
	}

	public override float GetFloat(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override Guid GetGuid(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override short GetInt16(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override long GetInt64(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override int GetOrdinal(string name)
	{
		throw new NotImplementedException();
	}

	public override int GetValues(object[] values)
	{
		throw new NotImplementedException();
	}

	public override bool NextResult()
	{
		return false;
	}

	#endregion
}