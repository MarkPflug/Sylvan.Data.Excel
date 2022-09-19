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
	protected abstract string GetFile([CallerMemberName] string name = "");

	public abstract ExcelWorkbookType WorkbookType { get; }
	public object Enumable { get; private set; }

	static void Unpack(string file, string folder)
	{
		// useful for debugging.
		try
		{
			Directory.Delete("unpack", true);
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
	public void TestCommonTypes()
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
			w.Write("data", reader);
		}
	}

	[Fact]
	public void WriteBigString()
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
			w.Write("data", reader);
		}
		Open(f);
	}

	[Fact]
	public void WriteNullCharString()
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
		var w = ExcelDataWriter.Create(f);
		w.Write("data", reader);
		// TODO: How expensive would it be to detect this in Write?
		Assert.ThrowsAny<Exception>(() => w.Dispose());
	}

	[Fact]
	public void WriteBoolean()
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
			w.Write("data", reader);
		}
		Open(f);
	}

#if NET6_0_OR_GREATER

	[Fact]
	public void TestDateOnly()
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
			w.Write("data", reader);
		}
	}

#endif
}
