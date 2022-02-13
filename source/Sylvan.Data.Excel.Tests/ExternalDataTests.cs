using System;
using System.Collections.Generic;
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
	public void GetExcelValues(string path)
	{
		if (path == null) return;
		var opts = new ExcelDataReaderOptions
		{
			Schema = ExcelSchema.NoHeaders
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
