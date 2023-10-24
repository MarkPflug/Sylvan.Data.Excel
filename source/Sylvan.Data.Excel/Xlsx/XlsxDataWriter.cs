using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace Sylvan.Data.Excel.Xlsx;

sealed partial class XlsxDataWriter : ExcelDataWriter
{
	sealed class SharedStringTable
	{
		readonly Dictionary<SharedStringEntry, string> dict;
		readonly List<SharedStringEntry> entries;

		public int UniqueCount => entries.Count;

		public string this[int idx] => entries[idx].str;

		public SharedStringTable()
		{
			const int InitialSize = 128;
			this.dict = new Dictionary<SharedStringEntry, string>(InitialSize);
			this.entries = new List<SharedStringEntry>(InitialSize);
		}

		struct SharedStringEntry : IEquatable<SharedStringEntry>
		{
			public string str;
			public string idxStr;

			public SharedStringEntry(string str)
			{
				this.str = str;
				this.idxStr = "";
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

		public string GetString(string str)
		{
			var entry = new SharedStringEntry(str);
			string? idxStr;
			if (!dict.TryGetValue(entry, out idxStr))
			{
				idxStr = this.entries.Count.ToString();
				this.entries.Add(entry);
				this.dict.Add(entry, idxStr);
			}
			return idxStr;
		}
	}

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

	public XlsxDataWriter(Stream stream, ExcelDataWriterOptions options) : base(stream, options)
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

		var fieldWriters = new FieldWriter[data.FieldCount];
		for (int i = 0; i < fieldWriters.Length; i++)
		{
			fieldWriters[i] = FieldWriter.Get(data.GetFieldType(i));
		}

		this.worksheets.Add(worksheetName);
		var idx = this.worksheets.Count;
		var entryName = "xl/worksheets/sheet" + idx + ".xml";
		var entry = zipArchive.CreateEntry(entryName, compression);
		using var es = entry.Open();
		using var xw = new StreamWriter(es, Encoding.UTF8, 0x4000);
		xw.Write($"<worksheet xmlns=\"{NS}\">");
		// freeze the header row.
		xw.Write("<sheetViews><sheetView workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" state=\"frozen\"/></sheetView></sheetViews>");
		xw.Write("<cols>");
		for (int i = 0; i < fieldWriters.Length; i++)
		{
			var num = (i + 1).ToString();
			var width = fieldWriters[i].GetWidth(data, i);
			xw.Write($"<col min=\"{num}\" max=\"{num}\" width=\"{width}\"/>");
		}

		xw.Write("</cols>");
		xw.Write("<sheetData>");

		var context = new Context(this, xw, data);

		var row = 0;
		// headers
		{
			xw.Write("<row>");
			for (int i = 0; i < data.FieldCount; i++)
			{
				var colName = data.GetName(i);
				if (string.IsNullOrEmpty(colName))
				{
					xw.Write("<c/>");
				}
				else
				{
					xw.Write("<c t=\"s\"><v>");

					var ssIdx = this.sharedStrings.GetString(colName);
					xw.Write(ssIdx);

					xw.Write("</v></c>");
				}
			}

			xw.Write("</row>");
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

			xw.Write("<row>");
			var c = data.FieldCount;
			for (int i = 0; i < c; i++)
			{
				if (data.IsDBNull(i))
				{
					xw.Write("<c/>");
				}
				else
				{
					var fw = i < fieldWriters.Length ? fieldWriters[i] : FieldWriter.Object;
					fw.WriteField(context, i);
				}
			}
			xw.Write("</row>");
			row++;
			if (row >= 0x100000)
			{
				// avoid calling Read again so the reader will remain in a state
				// where it can be written to a different worksheet.
				complete = false;
				break;
			}
		}

		xw.Write("</sheetData>");

		var end = ExcelSchema.GetExcelColumnName(fieldWriters.Length - 1);
		
		if (this.autoFilterOnHeader)
		{
			// apply filter to header row
			xw.Write($"<autoFilter ref=\"A1:{end}{row}\"/>");	
		}
		
		xw.Write("</worksheet>");
		return new WriteResult(row, complete);
	}

	const string NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
	const string PkgRelNS = "http://schemas.openxmlformats.org/package/2006/relationships";
	const string ODRelNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
	
	const string CoreNS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
	const string ContentTypeNS = "http://schemas.openxmlformats.org/package/2006/content-types";


	const string WorkbookPath = "xl/workbook.xml";
	const string AppPath = "docProps/app.xml";

	void WriteSharedStrings()
	{
		var e = this.zipArchive.CreateEntry("xl/sharedStrings.xml", compression);
		using var s = e.Open();
		using var xw = XmlWriter.Create(s, OpenPackaging.XmlSettings);
		xw.WriteStartElement("sst", NS);
		xw.WriteStartAttribute("uniqueCount");
		var c = this.sharedStrings.UniqueCount;
		xw.WriteValue(c);
		xw.WriteEndAttribute();
		for (int i = 0; i < c; i++)
		{
			xw.WriteStartElement("si");
			xw.WriteStartElement("t");
			var str = this.sharedStrings[i];

			var encodedStr = OpenXmlCodec.EncodeString(str);
			if (HasWhiteSpace(encodedStr))
			{
				xw.WriteAttributeString("xml", "space", null, "preserve");
			}
			// using XmlWriter, so escaping xml-layer characters
			// is handled by the XmlWriter.
			xw.WriteValue(encodedStr);
			xw.WriteEndElement();
			xw.WriteEndElement();
		}
		xw.WriteEndElement();
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
		var ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
		var wbName = "xl/workbook.xml";
		var e = this.zipArchive.CreateEntry(wbName, compression);

		using var s = e.Open();
		using var xw = XmlWriter.Create(s, OpenPackaging.XmlSettings);

		xw.WriteStartElement("workbook", ns);
		xw.WriteAttributeString("xmlns", "r", null, ODRelNS);

		xw.WriteStartElement("sheets", ns);
		for (int i = 0; i < this.worksheets.Count; i++)
		{
			var num = i + 1;
			xw.WriteStartElement("sheet", ns);

			xw.WriteStartAttribute("name");
			xw.WriteValue(this.worksheets[i]);
			xw.WriteEndAttribute();

			xw.WriteStartAttribute("sheetId");
			xw.WriteValue(num);
			xw.WriteEndAttribute();

			xw.WriteStartAttribute("id", ODRelNS);
			xw.WriteValue("s" + num);
			xw.WriteEndAttribute();

			xw.WriteEndElement();
		}
		xw.WriteEndElement();
		xw.WriteEndElement();
	}

	void WriteCoreProps()
	{
		var appEntry = zipArchive.CreateEntry("docProps/core.xml", compression);
		using var appStream = appEntry.Open();
		using var xw = XmlWriter.Create(appStream, OpenPackaging.XmlSettings);
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
			var entry = zipArchive.CreateEntry("_rels/.rels", compression);
			using var appStream = entry.Open();
			using var xw = XmlWriter.Create(appStream, OpenPackaging.XmlSettings);
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
			var entry = zipArchive.CreateEntry("xl/_rels/workbook.xml.rels", compression);
			using var appStream = entry.Open();
			using var xw = XmlWriter.Create(appStream, OpenPackaging.XmlSettings);
			xw.WriteStartElement("Relationships", PkgRelNS);

			xw.WriteStartElement("Relationship", PkgRelNS);
			xw.WriteAttributeString("Id", "s");
			xw.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
			xw.WriteAttributeString("Target", "styles.xml");
			xw.WriteEndElement();

			xw.WriteStartElement("Relationship", PkgRelNS);
			xw.WriteAttributeString("Id", "ss");
			xw.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
			xw.WriteAttributeString("Target", "sharedStrings.xml");
			xw.WriteEndElement();

			for (int i = 0; i < worksheets.Count; i++)
			{
				var num = (i + 1).ToString();
				xw.WriteStartElement("Relationship", PkgRelNS);
				xw.WriteAttributeString("Id", "s" + num);
				xw.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
				xw.WriteAttributeString("Target", "worksheets/sheet" + num + ".xml");
				xw.WriteEndElement();
			}

			xw.WriteEndElement();
		}

		// content types
		{
			var entry = zipArchive.CreateEntry("[Content_Types].xml", compression);
			using var appStream = entry.Open();
			using var xw = XmlWriter.Create(appStream, OpenPackaging.XmlSettings);
			xw.WriteStartElement("Types", ContentTypeNS);

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
			Override(xw, "/xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
			Override(xw, "/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
			Override(xw, "/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
			Override(xw, "/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml");

			for (int i = 0; i < worksheets.Count; i++)
			{
				var num = (i + 1).ToString();
				Override(xw, "/xl/worksheets/sheet" + num + ".xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
			}

			xw.WriteEndElement();
		}
	}

	void WriteStyles()
	{
		var appEntry = zipArchive.CreateEntry("xl/styles.xml", compression);
		using var appStream = appEntry.Open();
		using var wx = XmlWriter.Create(appStream, OpenPackaging.XmlSettings);
		wx.WriteStartElement("styleSheet", NS);

		wx.WriteStartElement("numFmts", NS);
		for (int i = 0; i < Formats.Length; i++)
		{
			wx.WriteStartElement("numFmt", NS);

			wx.WriteStartAttribute("numFmtId");
			wx.WriteValue(FormatOffset + i);
			wx.WriteEndAttribute();

			wx.WriteStartAttribute("formatCode");
			wx.WriteValue(Formats[i]);
			wx.WriteEndAttribute();
			wx.WriteEndElement();
		}

		wx.WriteEndElement();

		wx.WriteStartElement("fonts", NS);
		wx.WriteStartElement("font", NS);
		wx.WriteStartElement("name", NS);
		wx.WriteAttributeString("val", "Calibri");
		wx.WriteEndElement();
		wx.WriteEndElement();
		wx.WriteEndElement();

		wx.WriteStartElement("fills");
		wx.WriteStartElement("fill");
		wx.WriteEndElement();
		wx.WriteEndElement();

		wx.WriteStartElement("borders");
		wx.WriteStartElement("border");
		wx.WriteEndElement();
		wx.WriteEndElement();

		wx.WriteStartElement("cellStyleXfs");
		wx.WriteStartElement("xf");
		wx.WriteEndElement();
		wx.WriteEndElement();

		wx.WriteStartElement("cellXfs", NS);

		{
			wx.WriteStartElement("xf", NS);

			wx.WriteStartAttribute("numFmtId");
			wx.WriteValue(0);
			wx.WriteEndAttribute();

			wx.WriteStartAttribute("xfId");
			wx.WriteValue(0);
			wx.WriteEndAttribute();

			wx.WriteEndElement();
		}

		for (int i = 0; i < Formats.Length; i++)
		{
			wx.WriteStartElement("xf", NS);

			wx.WriteStartAttribute("numFmtId");
			wx.WriteValue(FormatOffset + i);
			wx.WriteEndAttribute();

			wx.WriteStartAttribute("xfId");
			wx.WriteValue(0);
			wx.WriteEndAttribute();

			wx.WriteEndElement();
		}

		wx.WriteEndElement();

		wx.WriteStartElement("cellStyles");
		wx.WriteStartElement("cellStyle");
		wx.WriteAttributeString("name", "Normal");
		wx.WriteAttributeString("xfId", "0");
		wx.WriteEndElement();
		wx.WriteEndElement();

		wx.WriteEndElement();
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
