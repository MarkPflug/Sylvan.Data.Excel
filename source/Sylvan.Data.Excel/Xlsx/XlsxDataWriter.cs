using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace Sylvan.Data.Excel.Xlsx;

sealed partial class XlsxDataWriter : ExcelDataWriter
{
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

	ZipArchive zipArchive;
	List<string> worksheets;
	List<string> formats = new List<string>();
	const CompressionLevel Compression = CompressionLevel.Optimal;
	bool truncateStrings;

	public XlsxDataWriter(Stream stream, ExcelDataWriterOptions options) : base(stream, options)
	{
		this.zipArchive = new ZipArchive(stream, ZipArchiveMode.Create, true);

		this.worksheets = new List<string>();
		this.formats = new List<string>();
		// used for datetime
		this.formats.Add("yyyy\\-mm\\-dd\\ hh:mm:ss.000");
		// used for dateonly
		this.formats.Add("yyyy\\-mm\\-dd");
		// used for timeonly
		this.formats.Add("hh:mm:ss");
		this.truncateStrings = options.TruncateStrings;
	}

	public override WriteResult Write(DbDataReader data, string? worksheetName)
	{
		return WriteInternal(data, worksheetName, false, default).GetAwaiter().GetResult();
	}

	public override async Task<WriteResult> WriteAsync(DbDataReader data, string? worksheetName, CancellationToken cancel)
	{
		return await WriteInternal(data, worksheetName, true, default);
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
		var entry = zipArchive.CreateEntry(entryName, Compression);
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
				if (!await data.ReadAsync(cancel))
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
				var fw = i < fieldWriters.Length ? fieldWriters[i] : ObjectFieldWriter.Instance;
				if (data.IsDBNull(i))
				{
					xw.Write("<c/>");
				}
				else
				{
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
		// apply filter to header row
		xw.Write($"<autoFilter ref=\"A1:{end}{row}\"/>");
		xw.Write("</worksheet>");
		return new WriteResult(row, complete);
	}

	const string NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
	const string PkgRelNS = "http://schemas.openxmlformats.org/package/2006/relationships";
	const string ODRelNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
	const string PropNS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
	const string CoreNS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
	const string ContentTypeNS = "http://schemas.openxmlformats.org/package/2006/content-types";


	const string WorkbookPath = "xl/workbook.xml";
	const string AppPath = "docProps/app.xml";

	void WriteSharedStrings()
	{
		var e = this.zipArchive.CreateEntry("xl/sharedStrings.xml", Compression);
		using var s = e.Open();
		using var xw = XmlWriter.Create(s, XmlSettings);
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
		var e = this.zipArchive.CreateEntry(wbName, Compression);

		using var s = e.Open();
		using var xw = XmlWriter.Create(s, XmlSettings);

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
			xw.WriteAttributeString("Id", "wb");
			xw.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
			xw.WriteAttributeString("Target", AppPath);

			xw.WriteEndElement();

			xw.WriteStartElement("Relationship", PkgRelNS);
			xw.WriteAttributeString("Id", "app");
			xw.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
			xw.WriteAttributeString("Target", WorkbookPath);

			xw.WriteEndElement();

			xw.WriteEndElement();
		}

		// workbook rels
		{
			var entry = zipArchive.CreateEntry("xl/_rels/workbook.xml.rels", Compression);
			using var appStream = entry.Open();
			using var xw = XmlWriter.Create(appStream, XmlSettings);
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
			var entry = zipArchive.CreateEntry("[Content_Types].xml", Compression);
			using var appStream = entry.Open();
			using var xw = XmlWriter.Create(appStream, XmlSettings);
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
		var appEntry = zipArchive.CreateEntry("xl/styles.xml", Compression);
		using var appStream = appEntry.Open();
		using var wx = XmlWriter.Create(appStream, XmlSettings);
		wx.WriteStartElement("styleSheet", NS);

		wx.WriteStartElement("numFmts", NS);
		for (int i = 0; i < formats.Count; i++)
		{
			wx.WriteStartElement("numFmt", NS);

			wx.WriteStartAttribute("numFmtId");
			wx.WriteValue(FormatOffset + i);
			wx.WriteEndAttribute();

			wx.WriteStartAttribute("formatCode");
			wx.WriteValue(formats[i]);
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
		//appX.WriteStartAttribute("count");
		//appX.WriteValue(formats.Count + 1);
		//appX.WriteEndAttribute();

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

		for (int i = 0; i < formats.Count; i++)
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
