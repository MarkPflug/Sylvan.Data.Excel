using Sylvan.Data.Csv;
using System;
using System.Data.SqlClient;
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
		var data = new[]
		{
			new {
				Name = "A",
				Data = "Alphabet".ToCharArray(),
			},
			new {
				Name = "B",
				Data = new string('Z',short.MaxValue).ToCharArray(),
			}
		};

		var dr = data.AsDataReader();
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
			Assert.Equal("A", edr.GetString(0));
			Assert.Equal("Alphabet", edr.GetString(1));
			Assert.True(edr.Read());
			Assert.Equal("B", edr.GetString(0));
			Assert.Equal(new string('Z', short.MaxValue), edr.GetString(1));
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
