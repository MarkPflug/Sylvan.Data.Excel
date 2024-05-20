#if NET6_0_OR_GREATER

using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace Sylvan.Data.Excel.Xlsb;

static class XlsbWriterExtensions
{
	public static void WriteType(this BinaryWriter bw, RecordType type)
	{
		var val = (int)type;
		if (val < 0x80)
		{
			bw.Write((byte)val); return;
		}
		else
		{
			bw.Write((byte)(0x80 | val & 0x7f));
			bw.Write((byte)(val >> 7));
		}
	}

	public static void WriteXF(this BinaryWriter bw, int fmtId, bool style = false)
	{
		bw.WriteType(RecordType.XF);
		bw.Write7BitEncodedInt(16);
		bw.Write((short)(style ? 0xffff : 0));//parent
		bw.Write((short)fmtId);//fmt
		bw.Write(0); // font, fill
		bw.Write(0); // border, rotation, indent
		bw.Write((short)0); // flags
		bw.Write((byte)1); // flag to apply format
		bw.Write((byte)0); //unused
	}

	public static void WriteRow(this BinaryWriter bw, int idx, int fieldCount)
	{
		Span<byte> data = stackalloc byte[27];
		data[0] = (byte)RecordType.Row;
		data[1] = 25; // len
		BitConverter.TryWriteBytes(data.Slice(2), idx);
		BitConverter.TryWriteBytes(data.Slice(15), 1);
		BitConverter.TryWriteBytes(data.Slice(23), fieldCount);
		bw.Write(data);
	}

	public static void WriteBlankCell(this BinaryWriter bw, int col)
	{
		// Write ROW
		bw.WriteType(RecordType.CellBlank);
		// len
		bw.Write7BitEncodedInt(8);
		// row
		bw.Write(col);
		// sf
		bw.Write(0);
	}

	public static void WriteSharedString(this BinaryWriter bw, int col, int ssIdx)
	{
		Span<byte> data = stackalloc byte[14];
		data[0] = (byte)RecordType.CellIsst;
		data[1] = 12; //len
		BitConverter.TryWriteBytes(data.Slice(2), col);
		BitConverter.TryWriteBytes(data.Slice(10), ssIdx);
		bw.Write(data);
	}

	static void WriteRK(this BinaryWriter bw, int col, uint rk, int fmt)
	{
		Span<byte> data = stackalloc byte[14];
		data[0] = (byte)RecordType.CellRK;
		data[1] = 12; //len
		BitConverter.TryWriteBytes(data.Slice(2), col);
		BitConverter.TryWriteBytes(data.Slice(6), fmt);
		BitConverter.TryWriteBytes(data.Slice(10), rk);
		bw.Write(data);
	}

	public static void WriteNumber(this BinaryWriter bw, int col, int val, int fmt = 0)
	{
		if (val >= 0 && val < 0x1fffffff)
		{
			var rk = (uint)(val << 2) | 0x2;
			WriteRK(bw, col, rk, fmt);
		}
		else
		{
			WriteNumber(bw, col, (double)val, fmt);
		}
	}

	public static void WriteNumber(this BinaryWriter bw, int col, double value, int fmt = 0)
	{
		var l = BitConverter.DoubleToInt64Bits(value);
		// write the value as an RK value if it can be done losslessly.
		if (((uint)l & 0xffffffff) == 0)
		{
			var rk = (uint)(l >> 32) & 0xfffffffc;
			WriteRK(bw, col, rk, fmt);
		}
		else
		{
			Span<byte> data = stackalloc byte[18];
			data[0] = (byte)RecordType.CellReal;
			data[1] = 16; //len
			BitConverter.TryWriteBytes(data.Slice(2), col);
			BitConverter.TryWriteBytes(data.Slice(6), fmt);
			BitConverter.TryWriteBytes(data.Slice(10), value);
			bw.Write(data);
		}
	}

	const int MaxRKInt = 0x1fffffff;
	const decimal MaxRKDec = MaxRKInt;

	public static void WriteNumber(this BinaryWriter bw, int col, decimal val, int fmt = 0)
	{
		if (val >= 0m && val <= MaxRKDec)
		{
			var i = (int)val;
			if (i == val)
			{
				// the value can be written as an integer
				bw.WriteNumber(col, i, fmt);
				return;
			}
			else
			{
				var mul = val * 100m;
				if (mul <= MaxRKDec)
				{
					var imul = (int)mul;
					if (mul == imul && imul >= 0 && imul <= MaxRKInt)
					{
						// the value can be written as a 100 scaled rk int
						var rk = (uint)(imul << 2) | 0x3;
						WriteRK(bw, col, rk, fmt);
						return;
					}
				}
			}
		}
		// TODO: I worry there is loss of precision with this cast.
		// should there be an option to throw when the double value doesn't roundtrip?
		bw.WriteNumber(col, (double)val, fmt);
	}

	public static void WriteBool(this BinaryWriter bw, int col, bool value)
	{
		// Write ROW
		bw.WriteType(RecordType.CellBool);
		// len
		bw.Write7BitEncodedInt(9);
		// row
		bw.Write(col);
		// sf
		bw.Write(0);

		bw.Write(value ? (byte)1 : (byte)0);
	}


	public static void WriteFrozenPane(this BinaryWriter bw)
	{
		// Write ROW
		bw.WriteType(RecordType.Pane);
		// len
		bw.Write7BitEncodedInt(29);

		bw.Write((double)0);

		bw.Write((double)1);

		// B1 active cell
		bw.Write(1);
		bw.Write(0);
		// bottom-left pane active
		bw.Write(2);

		bw.Write((byte)1);
	}

	public static void WriteViewStart(this BinaryWriter bw)
	{
		// Write ROW
		bw.WriteType(RecordType.WsViewStart);
		// len
		bw.Write7BitEncodedInt(30);

		bw.Write((short)0x03dc); // flags
		bw.Write(0); // normal view
		bw.Write(0); // top
		bw.Write(0); // left

		bw.Write((byte)4); // blue lines?
		bw.Write((byte)0);
		bw.Write((short)0);
		bw.Write((short)100); //scale
		bw.Write((short)0); //scale norm
		bw.Write((short)0); //scale SLV
		bw.Write((short)0); //scale PLV

		bw.Write(0); // wbk view
	}

	public static void WriteColWidth(this BinaryWriter bw, int col, double width)
	{
		// Write ROW
		bw.WriteType(RecordType.ColInfo);
		// len
		bw.Write7BitEncodedInt(18);

		bw.Write(col); // first
		bw.Write(col); // last
		var wx = (int)(width * 255);
		wx = Math.Min(wx, ushort.MaxValue);
		bw.Write(wx);
		bw.Write(0); //ixfe = default
		bw.Write((short)0x0002); // flags = user-width
	}

	public static void WriteWorksheetStart(this BinaryWriter bw)
	{
		bw.WriteMarker(RecordType.SheetStart);
	}

	public static void WriteWorksheetEnd(this BinaryWriter bw)
	{
		bw.WriteMarker(RecordType.SheetEnd);
	}

	public static void WriteWorkbookStart(this BinaryWriter bw)
	{
		bw.WriteMarker(RecordType.BookBegin);
	}

	public static void WriteWorkbookEnd(this BinaryWriter bw)
	{
		bw.WriteMarker(RecordType.BookEnd);
	}

	public static void WriteBundleStart(this BinaryWriter bw)
	{
		bw.WriteMarker(RecordType.BundleBegin);
	}

	public static void WriteBundleSheet(this BinaryWriter bw, int idx, string name)
	{
		bw.WriteType(RecordType.BundleSheet);
		var id = idx + 1;

		var relId = "s" + id;

		var len =
			8 +
			4 + (relId.Length * 2) +
			4 + (name.Length * 2);

		bw.Write7BitEncodedInt(len);
		bw.Write(0); // state (vis)
		bw.Write(id); // id
		bw.WriteString(relId);
		bw.WriteString(name);
	}

	public static void WriteString(this BinaryWriter bw, string str)
	{
		bw.Write(str.Length);
		var bs = MemoryMarshal.Cast<char, byte>(str.AsSpan());
		bw.Write(bs);
	}

	public static void WriteBundleEnd(this BinaryWriter bw)
	{
		bw.WriteMarker(RecordType.BundleEnd);
	}

	public static void WriteMarker(this BinaryWriter bw, RecordType type)
	{
		bw.WriteType(type);
		bw.Write((byte)0);
	}

	public static void WriteFont(this BinaryWriter bw, string name)
	{
		bw.WriteType(RecordType.Font);

		var len = 21 + (4 + 2 * name.Length);
		bw.Write7BitEncodedInt(len);
		bw.Write((short)0xdc);// height
		bw.Write((short)0); //grbit
		bw.Write((short)0x190);// weight
		bw.Write((short)0); //sss
		bw.Write((byte)0);// underline
		bw.Write((byte)2);// style = swiss
		bw.Write((byte)0);// charset
		bw.Write((byte)0);// unused
		bw.Write((long)0);// color = auto
		bw.Write((byte)0);// scheme
		bw.WriteString(name);
	}
}

sealed partial class XlsbDataWriter : ExcelDataWriter
{
	sealed class SharedStringTable
	{
		readonly Dictionary<SharedStringEntry, int> dict;
		readonly List<SharedStringEntry> entries;
		int count;
		public int UniqueCount => entries.Count;
		public int Count => count;

		public string this[int idx] => entries[idx].str;

		public SharedStringTable()
		{
			const int InitialSize = 128;
			this.dict = new Dictionary<SharedStringEntry, int>(InitialSize);
			this.entries = new List<SharedStringEntry>(InitialSize);
		}

		struct SharedStringEntry : IEquatable<SharedStringEntry>
		{
			public string str;
			public int idx;

			public SharedStringEntry(string str)
			{
				this.str = str;
				this.idx = 0;
			}

			public override int GetHashCode()
			{
				return str.GetHashCode();
			}

			public override bool Equals(object? obj)
			{
				return obj is SharedStringEntry e && this.Equals(e);
			}

			public bool Equals(SharedStringEntry other)
			{
				return this.str.Equals(other.str);
			}
		}

		public int GetString(string str)
		{
			var entry = new SharedStringEntry(str);
			int idx;
			this.count++;
			if (!dict.TryGetValue(entry, out idx))
			{
				idx = this.entries.Count;
				this.entries.Add(entry);
				this.dict.Add(entry, idx);
			}
			return idx;
		}
	}

	static readonly XmlWriterSettings XmlSettings =
	new XmlWriterSettings
	{
		OmitXmlDeclaration = true,
		CheckCharacters = false,
	};

	const int FormatOffset = 165;
	const int StringLimit = short.MaxValue;
	const int MaxWorksheetNameLength = 31;

	readonly ZipArchive zipArchive;
	readonly List<string> worksheets;
	readonly SharedStringTable sharedStrings;

	static string[] Formats = new[]
	{
		// used for datetime
		"yyyy\\-mm\\-dd\\ hh:mm:ss.000",
		// used for dateonly
		"yyyy\\-mm\\-dd",
		// used for timeonly/timespan
		"hh:mm:ss",
	};

	CompressionLevel compression;

	public XlsbDataWriter(Stream stream, ExcelDataWriterOptions options) : base(stream, options)
	{
		this.sharedStrings = new SharedStringTable();
		this.zipArchive = new ZipArchive(stream, ZipArchiveMode.Create, true);
		this.compression = options.CompressionLevel;
		this.worksheets = new List<string>();
	}

	public override WriteResult Write(DbDataReader data, string? worksheetName)
	{
		return WriteInternal(data, worksheetName, false, default).GetAwaiter().GetResult();
	}

	public override async Task<WriteResult> WriteAsync(DbDataReader data, string? worksheetName, CancellationToken cancel)
	{
		return await WriteInternal(data, worksheetName, true, default).ConfigureAwait(false);
	}

	async Task<WriteResult> WriteInternal(DbDataReader data, string? worksheetName, bool async, CancellationToken cancel)
	{
		if (worksheetName != null && worksheetName.Length > MaxWorksheetNameLength)
			throw new ArgumentException($"{nameof(worksheetName)}: must be under {MaxWorksheetNameLength} characters. ('{worksheetName}')");

		if (worksheetName != null && this.worksheets.Contains(worksheetName))
			throw new ArgumentException($"{nameof(worksheetName)}: worksheet '{worksheetName}' already exists.");

		if (worksheetName == null)
		{
			var sheetIdx = worksheets.Count;

			do
			{
				sheetIdx++;
				worksheetName = "Sheet " + sheetIdx;
			} while (worksheets.Contains(worksheetName));
		}

		this.worksheets.Add(worksheetName);
		var idx = this.worksheets.Count;
		var entryName = "xl/worksheets/sheet" + idx + ".bin";
		var entry = zipArchive.CreateEntry(entryName, compression);
		using var es = entry.Open();
		using var bs = new BufferedStream(es, 0x8000);
		using var bw = new BinaryWriter(bs);

		var fieldCount = data.FieldCount;

		var context = new Context(this, bw, data);

		var fieldWriters = new FieldWriter[fieldCount];
		for (int i = 0; i < fieldWriters.Length; i++)
		{
			fieldWriters[i] = FieldWriter.Get(data.GetFieldType(i));
		}

		bw.WriteWorksheetStart();

		// freeze the header row
		bw.WriteMarker(RecordType.WsViewsStart);
		bw.WriteViewStart();
		bw.WriteFrozenPane();
		bw.WriteMarker(RecordType.WsViewEnd);
		bw.WriteMarker(RecordType.WsViewsEnd);

		// set the column widths
		bw.WriteMarker(RecordType.ColInfoStart);
		for (int i = 0; i < fieldWriters.Length; i++)
		{
			var d = fieldWriters[i].GetWidth(data, i);
			bw.WriteColWidth(i, d);
		}
		bw.WriteMarker(RecordType.ColInfoEnd);

		bw.WriteMarker(RecordType.DataStart);
		// TODO: handle column widths based on fieldwriters.
		var row = 0;
		// headers
		{

			bw.WriteRow(row, fieldCount);
			for (int i = 0; i < fieldCount; i++)
			{
				var colName = data.GetName(i);
				if (string.IsNullOrEmpty(colName))
				{
					bw.WriteBlankCell(i);
				}
				else
				{
					var ssIdx = this.sharedStrings.GetString(colName);
					bw.WriteSharedString(i, ssIdx);
				}
			}
			row++;
		}
		bool complete = true;
		while (true)
		{
			if (async)
			{
				if (!await data.ReadAsync(cancel).ConfigureAwait(false))
				{
					break;
				}
			}
			else
			{
				if (!data.Read())
				{
					break;
				}
			}

			// use data.FieldCount here instead of local to
			// allow a data reader to present "jagged" data.
			var c = data.FieldCount;
			bw.WriteRow(row, c);
			for (int i = 0; i < c; i++)
			{
				if (data.IsDBNull(i))
				{
					bw.WriteBlankCell(i);
				}
				else
				{
					var fw = i < fieldWriters.Length ? fieldWriters[i] : FieldWriter.Object;
					fw.WriteField(context, i);
				}
			}
			row++;
			if (row >= 0x100000)
			{
				// avoid calling Read again so the reader will remain in a state
				// where it can be written to a different worksheet.
				complete = false;
				break;
			}
		}
		bw.WriteMarker(RecordType.DataEnd);

		// apply filtering to the header row.
		bw.WriteType(RecordType.FilterStart);
		bw.Write7BitEncodedInt(16);
		bw.Write(0);
		bw.Write(0);
		bw.Write(0);
		bw.Write(fieldCount - 1);
		bw.WriteMarker(RecordType.FilterEnd);

		// ignore number as string errors
		bw.WriteMarker(RecordType.BeginIgnoreError);

		bw.WriteType(RecordType.IgnoreError);
		bw.Write7BitEncodedInt(24);
		bw.Write(4); // 4 == ignore string as number
		bw.Write(1); // 1 range
		bw.Write(0); // row start
		bw.Write(row - 1); // row end
		bw.Write(0); // col start
		bw.Write(fieldCount - 1); // col end

		bw.WriteMarker(RecordType.EndIgnoreError);

		bw.WriteWorksheetEnd();
		return new WriteResult(row, complete);
	}

	const string PkgRelNS = "http://schemas.openxmlformats.org/package/2006/relationships";
	const string ContentTypeNS = "http://schemas.openxmlformats.org/package/2006/content-types";

	const string WorkbookPath = "xl/workbook.bin";

	void WriteSharedStrings()
	{
		var e = this.zipArchive.CreateEntry("xl/sharedStrings.bin", compression);
		using var s = e.Open();
		using var bw = new BinaryWriter(s);

		var c = this.sharedStrings.UniqueCount;
		bw.WriteType(RecordType.SSTBegin);
		bw.Write7BitEncodedInt(8);
		bw.Write(this.sharedStrings.Count);
		bw.Write(this.sharedStrings.UniqueCount);

		for (int i = 0; i < c; i++)
		{
			bw.WriteType(RecordType.SSTItem);

			var str = this.sharedStrings[i];
			var len = 1 + (4 + str.Length * 2);
			bw.Write7BitEncodedInt(len);
			bw.Write((byte)0);
			bw.WriteString(str);
		}
		bw.WriteMarker(RecordType.SSTEnd);
	}

	void WriteWorkbook()
	{
		var wbName = WorkbookPath;
		var e = this.zipArchive.CreateEntry(wbName, compression);

		using var s = e.Open();
		using var bw = new BinaryWriter(s);

		bw.WriteWorkbookStart();
		bw.WriteBundleStart();

		for (int i = 0; i < this.worksheets.Count; i++)
		{
			var num = i + 1;
			bw.WriteBundleSheet(i, this.worksheets[i]);
		}
		bw.WriteBundleEnd();
		bw.WriteWorkbookEnd();
	}

	void WritePkgMeta()
	{
		// pkg rels
		{
			var entry = zipArchive.CreateEntry("_rels/.rels", compression);
			using var appStream = entry.Open();
			using var xw = XmlWriter.Create(appStream, XmlSettings);
			xw.WriteStartElement("Relationships", PkgRelNS);

			xw.WriteStartElement("Relationship", PkgRelNS);
			xw.WriteAttributeString("Id", "app");
			xw.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
			xw.WriteAttributeString("Target", OpenPackaging.AppPath);

			xw.WriteEndElement();

			xw.WriteStartElement("Relationship", PkgRelNS);
			xw.WriteAttributeString("Id", "wb");
			xw.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
			xw.WriteAttributeString("Target", WorkbookPath);

			xw.WriteEndElement();

			xw.WriteEndElement();
		}

		// workbook rels
		{
			var entry = zipArchive.CreateEntry("xl/_rels/workbook.bin.rels", compression);
			using var appStream = entry.Open();
			using var xw = XmlWriter.Create(appStream, XmlSettings);
			xw.WriteStartElement("Relationships", PkgRelNS);

			xw.WriteStartElement("Relationship", PkgRelNS);
			xw.WriteAttributeString("Id", "s");
			xw.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
			xw.WriteAttributeString("Target", "styles.bin");
			xw.WriteEndElement();

			xw.WriteStartElement("Relationship", PkgRelNS);
			xw.WriteAttributeString("Id", "ss");
			xw.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
			xw.WriteAttributeString("Target", "sharedStrings.bin");
			xw.WriteEndElement();

			for (int i = 0; i < worksheets.Count; i++)
			{
				var num = (i + 1).ToString();
				xw.WriteStartElement("Relationship", PkgRelNS);
				xw.WriteAttributeString("Id", "s" + num);
				xw.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
				xw.WriteAttributeString("Target", "worksheets/sheet" + num + ".bin");
				xw.WriteEndElement();
			}

			xw.WriteEndElement();
		}

		// content types
		{
			var entry = zipArchive.CreateEntry("[Content_Types].xml", compression);
			using var appStream = entry.Open();
			using var xw = XmlWriter.Create(appStream, XmlSettings);
			xw.WriteStartElement("Types", ContentTypeNS);

			xw.WriteStartElement("Default", ContentTypeNS);
			xw.WriteAttributeString("Extension", "bin");
			xw.WriteAttributeString("ContentType", "application/vnd.ms-excel.sheet.binary.macroEnabled.main");
			xw.WriteEndElement();

			xw.WriteStartElement("Default", ContentTypeNS);
			xw.WriteAttributeString("Extension", "xml");
			xw.WriteAttributeString("ContentType", "application/xml");
			xw.WriteEndElement();


			xw.WriteStartElement("Default", ContentTypeNS);
			xw.WriteAttributeString("Extension", "rels");
			xw.WriteAttributeString("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
			xw.WriteEndElement();

			static void Override(XmlWriter xw, string path, string type)
			{
				xw.WriteStartElement("Override", ContentTypeNS);
				xw.WriteAttributeString("PartName", path);
				xw.WriteAttributeString("ContentType", type);
				xw.WriteEndElement();
			}

			Override(xw, "/xl/styles.bin", "application/vnd.ms-excel.styles");
			Override(xw, "/xl/sharedStrings.bin", "application/vnd.ms-excel.sharedStrings");
			Override(xw, "/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml");

			for (int i = 0; i < worksheets.Count; i++)
			{
				var num = (i + 1).ToString();
				Override(xw, "/xl/worksheets/sheet" + num + ".bin", "application/vnd.ms-excel.worksheet");
			}

			xw.WriteEndElement();
		}
	}

	void WriteStyles()
	{
		var styleEntry = zipArchive.CreateEntry("xl/styles.bin", compression);
		using var styleStream = styleEntry.Open();
		using var bs = new BufferedStream(styleStream, 0x4000);
		using var bw = new BinaryWriter(bs);

		bw.WriteMarker(RecordType.StyleBegin);

		bw.WriteType(RecordType.FmtStart);
		bw.Write7BitEncodedInt(4); // len
		bw.Write(Formats.Length);

		var idx = FormatOffset;
		foreach (var fmt in Formats)
		{
			bw.WriteType(RecordType.Fmt);
			var len = 2 + (4 + 2 * fmt.Length);
			bw.Write7BitEncodedInt(len);
			bw.Write((short)idx++);
			bw.WriteString(fmt);
		}

		bw.WriteMarker(RecordType.FmtEnd);

		bw.WriteType(RecordType.FontsStart);
		bw.Write7BitEncodedInt(4); // len
		bw.Write(1); // only 1 font
		bw.WriteFont("Calibri");
		bw.WriteMarker(RecordType.FontsEnd);

		bw.WriteType(RecordType.FillsStart);
		bw.Write7BitEncodedInt(4);
		bw.Write(1);

		bw.WriteType(RecordType.Fill);
		bw.Write7BitEncodedInt(68);
		bw.Write(new byte[] {
		  0x00, 0x00, 0x00, 0x00,
		  0x03, 0x40, 0x00, 0x00,
		  0x00, 0x00, 0x00, 0xff,
		  0x03, 0x41, 0x00, 0x00,
		  0xff, 0xff, 0xff, 0xff,
		  0x00, 0x00, 0x00, 0x00,
		  0x00, 0x00, 0x00, 0x00,
		  0x00, 0x00, 0x00, 0x00,
		  0x00, 0x00, 0x00, 0x00,
		  0x00, 0x00, 0x00, 0x00,
		  0x00, 0x00, 0x00, 0x00,
		  0x00, 0x00, 0x00, 0x00,
		  0x00, 0x00, 0x00, 0x00,
		  0x00, 0x00, 0x00, 0x00,
		  0x00, 0x00, 0x00, 0x00,
		  0x00, 0x00, 0x00, 0x00,
		  0x00, 0x00, 0x00, 0x00,
		});

		bw.WriteMarker(RecordType.FillsEnd);

		bw.WriteType(RecordType.BordersStart);
		bw.Write7BitEncodedInt(4);
		bw.Write(1);
		bw.WriteType(RecordType.Border);
		bw.Write7BitEncodedInt(51);
		bw.Write(new byte[]
		{
			  0x00, 0x00, 0x00, 0x01,
			  0x00, 0x00, 0x00, 0x00,
			  0x00, 0x00, 0x00, 0x00,
			  0x00, 0x01, 0x00, 0x00,
			  0x00, 0x00, 0x00, 0x00,
			  0x00, 0x00, 0x00, 0x01,
			  0x00, 0x00, 0x00, 0x00,
			  0x00, 0x00, 0x00, 0x00,
			  0x00, 0x01, 0x00, 0x00,
			  0x00, 0x00, 0x00, 0x00,
			  0x00, 0x00, 0x00, 0x01,
			  0x00, 0x00, 0x00, 0x00,
			  0x00, 0x00, 0x00,
		});

		bw.WriteMarker(RecordType.BordersEnd);

		bw.WriteType(RecordType.StyleXFsStart);
		bw.Write7BitEncodedInt(4);
		bw.Write(1);

		bw.WriteXF(0, true);
		bw.WriteMarker(RecordType.StyleXFsEnd);

		bw.WriteType(RecordType.CellXFStart);
		bw.Write7BitEncodedInt(4);
		bw.Write(Formats.Length + 1);

		bw.WriteXF(0);

		for (int i = 0; i < Formats.Length; i++)
		{
			bw.WriteXF(FormatOffset + i);
		}

		bw.WriteMarker(RecordType.CellXFEnd);

		bw.WriteMarker(RecordType.StyleEnd);
	}

	void Close()
	{
		OpenPackaging.WriteAppProps(this.zipArchive);
		WriteSharedStrings();
		WriteStyles();
		WriteWorkbook();
		WritePkgMeta();
	}

	public override void Dispose()
	{
		this.Close();
		this.zipArchive.Dispose();
		base.Dispose();
	}

#if ASYNC
	public override ValueTask DisposeAsync()
	{
		this.Close();
		this.zipArchive.Dispose();
		return base.DisposeAsync();
	}
#endif

}

#endif