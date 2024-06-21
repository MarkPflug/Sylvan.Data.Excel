#if NETCOREAPP3_1_OR_GREATER

using System;
using System.IO;
using Xunit;
using Xunit.Abstractions;

namespace Sylvan.Data.Excel;

public class OleStreamTests : ExternalDataTests
{
	public OleStreamTests(ITestOutputHelper o) : base(o)
	{
	}

	[Theory]
	[MemberData(nameof(GetXlsFiles))]
	public void Verify(string file)
	{
		file = GetFullPath(file);

		using var s = File.OpenRead(file);
		Ole2Package pkg;
		try
		{
			pkg = new Ole2Package(s);
		}
		catch (InvalidDataException)
		{
			return;
		}
		var e = pkg.GetEntry("Workbook\0");

		if (e == null)
		{
			return;
		}

		long len = 0;
		for (int i = 0; i < 3; i++)
		{
			using var es = e.Open();
			var x = ReadRandomly(es);
			if (len == 0)
			{
				len = x;
			}
			else
			{
				Assert.Equal(len, x);
			}
			TestOutput(x.ToString());
		}
	}

	static long ReadRandomly(Stream s)
	{
		var rand = new Random();
		var buffer = new byte[0x1000];
		long len = 0;
		while (true)
		{
			var readLen = rand.Next(600, buffer.Length);
			var l = s.Read(buffer, 0, readLen);
			len += l;
			if (l == 0)
				break;
		}
		return len;
	}
}


#endif