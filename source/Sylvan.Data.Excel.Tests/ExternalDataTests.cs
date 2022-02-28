using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Xunit;

namespace Sylvan.Data.Excel;

// Set the `SylvanExcelTestData` env var to point to a directory
// containing files that will be tested by this set of tests.
public class ExternalDataTests
{
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

		var offset = 0;
		while (offset < buf.Length)
		{
			var l = 999;// Random.Shared.Next(1000, 3000);
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
				Debug.WriteLine($"{p} {len}");
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
