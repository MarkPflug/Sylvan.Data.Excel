using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Xunit;
using Xunit.Abstractions;

namespace Sylvan.Data.Excel;

// Set the `SylvanExcelTestData` env var to point to a directory
// containing files that will be tested by this set of tests.
public class ExternalDataTests
{
	ITestOutputHelper o;

	public ExternalDataTests(ITestOutputHelper o)
	{
		this.o = o;
#if NETCOREAPP1_0_OR_GREATER
		Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
#endif
	}

	public static IEnumerable<object[]> GetInputs()
	{
		var paths = Environment.GetEnvironmentVariable("SylvanExcelTestData");
		if (string.IsNullOrEmpty(paths))
		{
			yield return new object[] { null };
			yield break;
		}

		foreach (var path in paths.Split(';'))
		{
			foreach (var file in Directory.EnumerateFiles(path, "*.xls*", SearchOption.TopDirectoryOnly))
			{
				yield return new object[] { file };
			}
		}
	}


	[Fact]
	public void AnalyzeFiles()
	{
		var root = Environment.GetEnvironmentVariable("SylvanExcelTestData");
		var files = Directory.EnumerateFiles(root, "*.xlsx");
		foreach (var file in files)
		{
			AnalyzeFile(file);
		}
	}

	void AnalyzeFile(string file)
	{
		try
		{
			//using var s = File.OpenRead(file);
			//using var za = new ZipArchive(s, ZipArchiveMode.Read);
			var edr = ExcelDataReader.Create(file);
			while (edr.Read())
			{
				for (int i = 0; i < edr.RowFieldCount; i++)
				{
					if (edr.GetExcelDataType(i) == ExcelDataType.String)
					{
						if (edr.GetString(i) == "")
						{
							o.WriteLine($"{Path.GetFileName(file)} {edr.RowNumber} {i}");
						}
					}
				}
			}
		}
		catch (Exception e)
		{
			o.WriteLine($"{Path.GetFileName(file)} ERROR {e.Message}");
		}
	}

	[Fact]
	public void XmlCharRegex()
	{
		var str = "ab\bcd";
		var rep = Regex.Replace(str, @"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "");
		Assert.Equal("abcd", rep);
	}

	[Theory]
	[MemberData(nameof(GetInputs))]
	public void ExtractWS(string path)
	{
		// this test was to debug the ole2stream buffer management
		if (path == null) return;
		var opts = new ExcelDataReaderOptions
		{
			Schema = ExcelSchema.NoHeaders
		};

		var stream = File.OpenRead(path);
		var pkg = new Ole2Package(stream);
		var part = pkg.GetEntry("Workbook\0");
		if (part == null)
			throw new InvalidDataException();
		var ps = part.Open();

		var buf = new byte[ps.Length];

		var rand = new Random();
		var offset = 0;
		while (offset < buf.Length)
		{
			var l = rand.Next(500, 1500);
			l = Math.Min(buf.Length - offset, l);
			var r = ps.Read(buf, offset, l);
			if (r == 0)
				break;
			offset += r;
		}

		var p = 0;
		var max = 0;
		while (p < buf.Length)
		{
			var code = BitConverter.ToUInt16(buf, p);
			p += 2;
			var len = BitConverter.ToUInt16(buf, p);
			if (len > max)
			{
				max = len;
				o.WriteLine($"{code} {p} {len}");
			}
			p += 2;
			p += len;

			//Debug.WriteLine($"{code} {len}");
		}
	}

	[Theory]
	[MemberData(nameof(GetInputs))]
	public void GetExcelValues(string path)
	{
		if (path == null) return;
		var opts = new ExcelDataReaderOptions
		{
			ReadHiddenWorksheets = true,
			Schema = ExcelSchema.NoHeaders,
		};

		var edr = ExcelDataReader.Create(path, opts);
		var filename = Path.GetFileName(path);

		do
		{
			var n = edr.WorksheetName;
			using var sw = File.CreateText(filename + "-" + n + ".txt");

			while (edr.Read())
			{
				for (int i = 0; i < edr.RowFieldCount; i++)
				{
					if (i > 0)
						sw.Write('\t');
					sw.Write(edr.GetExcelValue(i));
				}

				sw.WriteLine();
			}
			sw.Flush();
		} while (edr.NextResult());
	}

	[Theory]
	[MemberData(nameof(GetInputs))]
	public void GetValue(string path)
	{
		if (path == null) return;
		var opts = new ExcelDataReaderOptions
		{
			Schema = ExcelSchema.NoHeaders,
			GetErrorAsNull = true
		};
		var edr = ExcelDataReader.Create(path, opts);

		do
		{
			while (edr.Read())
			{
				for (int i = 0; i < edr.RowFieldCount; i++)
				{
					edr.GetValue(i);
				}
			}
		} while (edr.NextResult());
	}
}
