#if NET6_0_OR_GREATER

using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace Sylvan.Data.Excel.Xlsb;

static class XlsbWriterExtensions
{
	public static void WriteType(this BinaryWriter bw, RecordType type)
	{
		var val = (int)type;
		bw.Write7BitEncodedInt(val);
	}

	public static void WriteRow(this BinaryWriter bw, int idx)
	{
		// Write ROW
		bw.WriteType(RecordType.Row);
		// len
		bw.Write7BitEncodedInt(8);
		// row
		bw.Write(idx);
		// ifx
		bw.Write(0);
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
		// Write ROW
		bw.WriteType(RecordType.CellIsst);
		// len
		bw.Write7BitEncodedInt(12);
		// row
		bw.Write(col);
		// sf
		bw.Write(0);

		bw.Write(ssIdx);
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

	public static void WriteNumber(this BinaryWriter bw, int col, double value)
	{
		// Write ROW
		bw.WriteType(RecordType.CellNum);
		// len
		bw.Write7BitEncodedInt(12);
		// row
		bw.Write(col);
		// sf
		bw.Write(0);

		bw.Write(value);
	}

	public static void WriteBundleStart(this BinaryWriter bw)
	{
		bw.WriteType(RecordType.BundleBegin);
		bw.Write7BitEncodedInt(0);
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
		bw.Write(relId.Length);
		bw.Write(relId.AsSpan());
		bw.Write(name.Length);
		bw.Write(name.AsSpan());
	}

	public static void WriteBundleEnd(this BinaryWriter bw)
	{
		bw.WriteType(RecordType.BundleEnd);
		bw.Write7BitEncodedInt(0);
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
		//IndentChars = " ",
		//Indent = true,
		//NewLineChars = "\n",
		OmitXmlDeclaration = true,
		// We are handling this ourselves in the shared string handling.
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

	const CompressionLevel Compression = CompressionLevel.Optimal;

	public XlsbDataWriter(Stream stream, ExcelDataWriterOptions options) : base(stream, options)
	{
		this.sharedStrings = new SharedStringTable();
		this.zipArchive = new ZipArchive(stream, ZipArchiveMode.Create, true);
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
			throw new ArgumentException(nameof(worksheetName));

		if (worksheetName != null && this.worksheets.Contains(worksheetName))
			throw new ArgumentException(nameof(worksheetName));

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
		var entry = zipArchive.CreateEntry(entryName, Compression);
		using var es = entry.Open();
		using var bw = new BinaryWriter(es);

		var row = 0;
		// headers
		{
			bw.WriteRow(row);
			for (int i = 0; i < data.FieldCount; i++)
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

			bw.WriteRow(row);
			var c = data.FieldCount;
			for (int i = 0; i < c; i++)
			{				
				if (data.IsDBNull(i))
				{
					bw.WriteBlankCell(i);
				}
				else
				{
					var str = data.GetValue(i)?.ToString() ?? string.Empty;
					var ssIdx = this.sharedStrings.GetString(str);
					bw.WriteSharedString(i, ssIdx);
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

		return new WriteResult(row, complete);
	}

	const string NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
	const string PkgRelNS = "http://schemas.openxmlformats.org/package/2006/relationships";
	const string ODRelNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
	const string PropNS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
	const string CoreNS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
	const string ContentTypeNS = "http://schemas.openxmlformats.org/package/2006/content-types";


	const string WorkbookPath = "xl/workbook.bin";
	const string AppPath = "docProps/app.xml";

	void WriteSharedStrings()
	{
		var e = this.zipArchive.CreateEntry("xl/sharedStrings.bin", Compression);
		using var s = e.Open();
		using var bw = new BinaryWriter(s);


		var c = this.sharedStrings.UniqueCount;
		bw.WriteType(RecordType.SSTBegin);
		bw.Write7BitEncodedInt(8);
		// total count
		bw.Write(c);
		// count
		bw.Write(c);

		for (int i = 0; i < c; i++)
		{
			bw.WriteType(RecordType.SSTItem);

			var str = this.sharedStrings[i];
			var len = 1 + (4 + str.Length * 2);
			bw.Write((byte)0);
			bw.Write(str.Length);
			bw.Write(str.AsSpan());
		}
	}

	static bool HasWhiteSpace(string str)
	{
		char c;
		if (str.Length > 0)
		{
			c = str[0];
			if (char.IsWhiteSpace(c))
				return true;
			c = str[str.Length - 1];
			if (char.IsWhiteSpace(c))
				return true;
		}
		return false;
	}

	void WriteWorkbook()
	{
		var wbName = WorkbookPath;
		var e = this.zipArchive.CreateEntry(wbName, Compression);

		using var s = e.Open();
		using var bw = new BinaryWriter(s);

		bw.WriteBundleStart();
		
		for (int i = 0; i < this.worksheets.Count; i++)
		{
			var num = i + 1;
			bw.WriteBundleSheet(i, this.worksheets[i]);
		}
		bw.WriteBundleStart();
	}

	void WriteAppProps()
	{
		var appEntry = zipArchive.CreateEntry("docProps/app.xml", Compression);
		using var appStream = appEntry.Open();
		using var xw = XmlWriter.Create(appStream, XmlSettings);
		xw.WriteStartElement("Properties", PropNS);
		var asmName = Assembly.GetExecutingAssembly().GetName();
		xw.WriteStartElement("Application", PropNS);
		xw.WriteValue(asmName.Name);
		xw.WriteEndElement();
		xw.WriteStartElement("AppVersion", PropNS);
		var v = asmName.Version!;
		// AppVersion must be of the format XX.YYYY
		var ver = $"{v.Major:00}.{v.Minor:00}{v.Build:00}";
		xw.WriteValue(ver);
		xw.WriteEndElement();
		xw.WriteEndElement();
	}

	void WriteCoreProps()
	{
		var appEntry = zipArchive.CreateEntry("docProps/core.xml", Compression);
		using var appStream = appEntry.Open();
		using var xw = XmlWriter.Create(appStream, XmlSettings);
		xw.WriteStartElement("coreProperties", CoreNS);

		xw.WriteStartElement("lastModifiedBy", CoreNS);
		xw.WriteValue(Environment.UserName);
		xw.WriteEndElement();
		xw.WriteEndElement();
	}

	void WritePkgMeta()
	{
		// pkg rels
		{
			var entry = zipArchive.CreateEntry("_rels/.rels", Compression);
			using var appStream = entry.Open();
			using var xw = XmlWriter.Create(appStream, XmlSettings);
			xw.WriteStartElement("Relationships", PkgRelNS);

			xw.WriteStartElement("Relationship", PkgRelNS);
			xw.WriteAttributeString("Id", "app");
			xw.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
			xw.WriteAttributeString("Target", AppPath);

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
			var entry = zipArchive.CreateEntry("xl/_rels/workbook.bin.rels", Compression);
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
			var entry = zipArchive.CreateEntry("[Content_Types].xml", Compression);
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
		var styleEntry = zipArchive.CreateEntry("xl/styles.bin", Compression);
		using var styleStream = styleEntry.Open();
		using var s = typeof(XlsbDataWriter).Assembly.GetManifestResourceStream("styles");
		s!.CopyTo(styleStream);
		//using var bw = new BinaryWriter(appStream);

		//bw.WriteType(RecordType.StyleBegin);
		//bw.Write7BitEncodedInt(0);//len

		//bw.WriteType(RecordType.CellXFStart);
		//bw.Write7BitEncodedInt(4);
		//bw.write

		//for (int i = 0; i < Formats.Length; i++)
		//{
		//	wx.WriteStartElement("numFmt", NS);

		//	wx.WriteStartAttribute("numFmtId");
		//	wx.WriteValue(FormatOffset + i);
		//	wx.WriteEndAttribute();

		//	wx.WriteStartAttribute("formatCode");
		//	wx.WriteValue(Formats[i]);
		//	wx.WriteEndAttribute();
		//	wx.WriteEndElement();
		//}

		//wx.WriteEndElement();

		//wx.WriteStartElement("fonts", NS);
		//wx.WriteStartElement("font", NS);
		//wx.WriteStartElement("name", NS);
		//wx.WriteAttributeString("val", "Calibri");
		//wx.WriteEndElement();
		//wx.WriteEndElement();
		//wx.WriteEndElement();

		//wx.WriteStartElement("fills");
		//wx.WriteStartElement("fill");
		//wx.WriteEndElement();
		//wx.WriteEndElement();

		//wx.WriteStartElement("borders");
		//wx.WriteStartElement("border");
		//wx.WriteEndElement();
		//wx.WriteEndElement();

		//wx.WriteStartElement("cellStyleXfs");
		//wx.WriteStartElement("xf");
		//wx.WriteEndElement();
		//wx.WriteEndElement();

		//wx.WriteStartElement("cellXfs", NS);
		////appX.WriteStartAttribute("count");
		////appX.WriteValue(formats.Count + 1);
		////appX.WriteEndAttribute();

		//{
		//	wx.WriteStartElement("xf", NS);

		//	wx.WriteStartAttribute("numFmtId");
		//	wx.WriteValue(0);
		//	wx.WriteEndAttribute();

		//	wx.WriteStartAttribute("xfId");
		//	wx.WriteValue(0);
		//	wx.WriteEndAttribute();

		//	wx.WriteEndElement();
		//}

		//for (int i = 0; i < Formats.Length; i++)
		//{
		//	wx.WriteStartElement("xf", NS);

		//	wx.WriteStartAttribute("numFmtId");
		//	wx.WriteValue(FormatOffset + i);
		//	wx.WriteEndAttribute();

		//	wx.WriteStartAttribute("xfId");
		//	wx.WriteValue(0);
		//	wx.WriteEndAttribute();

		//	wx.WriteEndElement();
		//}

		//wx.WriteEndElement();

		//wx.WriteStartElement("cellStyles");
		//wx.WriteStartElement("cellStyle");
		//wx.WriteAttributeString("name", "Normal");
		//wx.WriteAttributeString("xfId", "0");
		//wx.WriteEndElement();
		//wx.WriteEndElement();

		//wx.WriteEndElement();
	}

	void Close()
	{
		// core.xml isn't needed.
		//WriteCoreProps();
		WriteAppProps();
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
}


#endif