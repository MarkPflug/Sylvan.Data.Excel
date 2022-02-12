using System;
using System.Collections.Generic;
using System.IO;
using Xunit;

namespace Sylvan.Data.Excel;

public class TestFile
{
	public TestFile(string path)
	{
		this.Name = System.IO.Path.GetFileName(path);
		this.Path = path;
	}

	public string Name { get; }
	public string Path { get; }

	public override string ToString()
	{
		return Name;
	}
}

// Set the `SylvanExcelTestData` env var to point to a directory
// containing files that will be tested by this set of tests.
public class ExternalDataTests
{
	public static IEnumerable<object[]> GetInputs()
	{
		var paths = Environment.GetEnvironmentVariable("SylvanExcelTestData");
		
		foreach(var path in paths.Split(';'))
		{
			foreach (var file in Directory.EnumerateFiles(path, "*.xls*", SearchOption.TopDirectoryOnly))
			{
				yield return new object[] { file };
			}
		}
	}	

	[Theory]
	[MemberData(nameof(GetInputs))]
	public void TestFile(string path)
	{
		var opts = new ExcelDataReaderOptions { 
			Schema = ExcelSchema.NoHeaders,
			GetErrorAsNull = true
		};
		var edr = ExcelDataReader.Create(path, opts);
		WriteSheets(Path.GetFileName(path), edr);
	}

	static void WriteSheets(string file, ExcelDataReader edr)
	{
		do
		{
			var n = edr.WorksheetName;
			using var sw = File.CreateText(file + "-" + n + ".txt");

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
	public void TestGetValues(string path)
	{
		var opts = new ExcelDataReaderOptions
		{
			Schema = ExcelSchema.NoHeaders,
			GetErrorAsNull = true
		};
		var edr = ExcelDataReader.Create(path, opts);
		WriteSheetValues(Path.GetFileName(path), edr);
	}

	static void WriteSheetValues(string file, ExcelDataReader edr)
	{
		do
		{
			var n = edr.WorksheetName;
			using var sw = File.CreateText(file + "-" + n + ".txt");

			while (edr.Read())
			{
				for (int i = 0; i < edr.RowFieldCount; i++)
				{
					if (i > 0)
						sw.Write('\t');
					sw.Write(edr.GetValue(i));
				}

				sw.WriteLine();
			}
			sw.Flush();
		} while (edr.NextResult());
	}
}
