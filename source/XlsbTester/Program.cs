using Sylvan.Data.Excel.Xlsb;
using System.IO.Compression;

var file = args.Length > 0 ? args[0] : "Test.xlsb";

Dump(file);

static void Dump(string file)
{
	var fs = File.OpenRead(file);

	var a = new ZipArchive(fs, ZipArchiveMode.Read);

	foreach (var e in a.Entries)
	{
		if (e.Name.EndsWith(".bin"))
		{
			if (!e.Name.Contains("sheet")) continue;

			Console.WriteLine(new string('-', 80));
			Console.WriteLine(e.Name);
			var s = e.Open();
			var r = new XlsbReader(s);
			while (r.ReadRecord())
			{
				Console.WriteLine(r.Type + " " + r.Length + " " + (RecordType)r.Type);
				if (r.Type == 0)
				{
					var dat = r.DataSpan;
				}

				if (r.Length > 0)
				{
					Console.Write("  ");
					switch (r.Type)
					{
						case 2: //RK
						case 5: // Real
						default:
							var data = r.DataSpan;
							int i = 0;
							// format the data so it aligns with
							// the documentation tables.
							foreach (var b in data)
							{
								Console.Write(b.ToString("x2"));
								Console.Write(' ');
								if (++i == 4)
								{
									i = 0;
									Console.WriteLine();
									Console.Write("  ");
								}
							}
							Console.WriteLine();
							break;
					}
				}
			}
		}
	}
}
