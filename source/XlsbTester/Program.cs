using Sylvan.Data.Excel;
using Sylvan.Data.Excel.Xlsb;
using System.IO.Compression;

//var name = "C:\\Users\\Mark\\source\\Sylvan.Data.Excel\\bin\\Debug\\net6.0\\simple.xlsb";
//var name = "/data/excel/input.xlsb";
var name = "/data/excel/fmt.xlsb";

//Dump(name);

var fs = File.OpenRead(name);

var a = new ZipArchive(fs, ZipArchiveMode.Read);

foreach (var e in a.Entries)
{
	if (e.Name.EndsWith(".bin"))
	{
		if (e.Name.Contains("printer")) continue;

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
							if(++i == 4)
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

#pragma warning disable CS8321 // Local function is declared but never used
static void Dump(string file)
{
	var r = ExcelDataReader.Create(file, new ExcelDataReaderOptions { Schema = ExcelSchema.NoHeaders });
	while (r.Read())
	{
		for (int i = 0; i < r.RowFieldCount; i++)
		{
			Console.WriteLine(r.GetDouble(i));
		}
		Console.WriteLine();
	}
}
#pragma warning restore CS8321 // Local function is declared but never used

sealed class XlsbReader
{
	byte[] data;
	int idx = 0;

	public XlsbReader(Stream stream)
	{
		var ms = new MemoryStream();
		stream.CopyTo(ms);
		data = ms.ToArray();
	}

	public ReadOnlySpan<byte> RecordSpan
	{
		get
		{
			return data.AsSpan(start, end - start);
		}
	}

	public ReadOnlySpan<byte> DataSpan
	{
		get
		{
			return data.AsSpan(dataStart, end - dataStart);
		}
	}

	public int Type => type;
	public int Length => len;

	int start;
	int dataStart;
	int end;
	int type;
	int len;

	public bool ReadRecord()
	{
		if (idx >= data.Length)
			return false;

		this.start = idx;

		var i = idx;

		this.type = ReadRecordType(ref i);
		this.len = ReadRecordLen(ref i);
		this.dataStart = i;

		i += len;

		this.end = i;
		this.idx = i;
		return true;
	}

	int ReadRecordType(ref int idx)
	{
		var b = data[idx++];
		int type;
		if (b >= 0x80)
		{
			var b2 = data[idx++];
			if (b2 >= 0x80)
				throw new InvalidDataException();
			type = (b & 0x7f) | (b2 << 7);
		}
		else
		{
			type = b;
		}
		return type;
	}

	int ReadRecordLen(ref int idx)
	{
		int accum = 0;
		int shift = 0;
		for (int i = 0; i < 4; i++, shift += 7)
		{
			var b = data[idx++];
			accum |= (b & 0x7f) << shift;
			if ((b & 0x80) == 0)
				break;
		}
		return accum;
	}
}


