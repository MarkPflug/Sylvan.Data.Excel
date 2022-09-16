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

	[Fact]
	public void Test1()
	{
		var f = GetFile();
		using var w = ExcelDataWriter.Create(f);

		var dat =
			Enumerable
			.Range(0, 100)
			.Select(i => new { Name = "n" + i, Id = i, Value = Math.PI * i })
			.AsDataReader();
		w.Write("data", dat);


	}

	[Fact]
	public void Test3()
	{
		const string query = @"

select
e.id,
e.Name, 
CreatedDate,
ClosedDate,
s.name as state,
o.name as Org,
coalesce(ca.Name, ca.firstname + ' ' + ca.lastname) as creator,
coalesce(ca.Name, ca.firstname + ' ' + ca.lastname) as owner

from sc.issue e
left join sc.issuestatus s
	on e.statusid  = s.id
left join core.Organization o
	on e.OrganizationId = o.Id
left join core.Account ca
	on e.CreatorId = ca.id
left join core.Account oa
	on e.OwnerId = oa.id";

		var f = GetFile();
		using (var w = ExcelDataWriter.Create(f))
		{
			var conn = new SqlConnection("Data Source=.;Initial Catalog=sc2;Integrated Security=true;");
			conn.Open();
			var cmd = conn.CreateCommand();
			cmd.CommandText = query;
			var data = cmd.ExecuteReader();
			w.Write("data", data);
		}

		try
		{

			Directory.Delete("unpack", true);
		}
		catch { }
		ZipFile.ExtractToDirectory(f, Path.GetDirectoryName(f) + "unpack");

		//var psi = new ProcessStartInfo(f)
		//{
		//	UseShellExecute = true,
		//};
		//Process.Start(psi);


	}

}
