using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.IO.Packaging;
using System.Reflection;
using System.Xml;

namespace Sylvan.Data.Excel.Xlsx;

sealed partial class XlsxDataWriter : ExcelDataWriter
{
	Package zipArchive;

	List<string> worksheets;

	int fmtOffset = 165;
	List<string> formats = new List<string>();
	CompressionOption compression = CompressionOption.Normal;

	public XlsxDataWriter(Stream stream) : base(stream)
	{
		this.zipArchive = Package.Open(stream, FileMode.CreateNew);
		
		this.worksheets = new List<string>();
		this.formats = new List<string>();
		this.formats.Add("yyyy\\-mm\\-dd\\ hh:mm:ss");
		this.formats.Add("yyyy\\-mm\\-dd\\ hh:mm:ss.000");
		this.formats.Add("yyyy\\-mm\\-dd");
	}

	public override void Write(string worksheetName, DbDataReader data)
	{
		if (this.worksheets.Contains(worksheetName))
			throw new ArgumentException(nameof(worksheetName));
		this.worksheets.Add(worksheetName);
		var idx = this.worksheets.Count;
		var partUri = new Uri("/xl/worksheets/sheet" + idx + ".xml", UriKind.Relative);
		var entry = zipArchive.CreatePart(partUri, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", compression);
		using var es = entry.GetStream();
		using var xw = XmlWriter.Create(es, new XmlWriterSettings { CheckCharacters = false });
		xw.WriteStartElement("worksheet", NS);
		xw.WriteStartElement("sheetData", NS);
		int row = 0;

		var context = new Context(this, xw, data);

		FieldWriter[] fieldWriter = new FieldWriter[data.FieldCount];
		for (int i = 0; i < fieldWriter.Length; i++)
		{
			fieldWriter[i] = FieldWriter.Get(data.GetFieldType(i));
		}

		// headers
		{
			row++;
			xw.WriteStartElement("row", NS);
			int last = -1;
			for (int i = 0; i < data.FieldCount; i++)
			{
				var colName = data.GetName(i);
				if (string.IsNullOrEmpty(colName)) { continue; }

				xw.WriteStartElement("c", NS);

				if (i != last + 1)
				{
					xw.WriteStartAttribute("r");
					var cn = ExcelSchema.GetExcelColumnName(i);
					xw.WriteValue(cn + "" + row);
					xw.WriteEndAttribute();
				}
				last = i;

				xw.WriteStartAttribute("t");
				xw.WriteValue("s");
				xw.WriteEndAttribute();

				xw.WriteStartElement("v", NS);

				var ssIdx = this.sharedStrings.GetString(colName);
				xw.WriteValue(ssIdx);

				xw.WriteEndElement();
				xw.WriteEndElement();
			}

			xw.WriteEndElement();
		}

		while (data.Read())
		{
			xw.WriteStartElement("row", NS);
			for (int i = 0; i < fieldWriter.Length; i++)
			{
				fieldWriter[i].WriteField(context, i);
			}
			xw.WriteEndElement();
		}

		xw.WriteEndElement();
		xw.WriteEndElement();
	}

	const string NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
	const string RelNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
	const string PropNS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
	const string CoreNS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";

	void WriteSharedStrings()
	{
		var e = this.zipArchive.CreatePart(new Uri("/xl/sharedStrings.xml", UriKind.Relative), "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", compression);
		using var s = e.GetStream();
		using var w = XmlWriter.Create(s);
		w.WriteStartElement("sst", NS);
		w.WriteStartAttribute("uniqueCount");
		var c = this.sharedStrings.UniqueCount;
		w.WriteValue(c);
		w.WriteEndAttribute();
		for (int i = 0; i < c; i++)
		{
			w.WriteStartElement("si");
			w.WriteStartElement("t");
			w.WriteValue(this.sharedStrings[i]);
			w.WriteEndElement();
			w.WriteEndElement();
		}
		w.WriteEndElement();
	}

	void WriteWorkbook()
	{
		var ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
		var wbUri = new Uri("/xl/workbook.xml", UriKind.Relative);
		var e = this.zipArchive.CreatePart(wbUri, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
		e.CreateRelationship(new Uri("/xl/sharedStrings.xml", UriKind.Relative), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
		e.CreateRelationship(new Uri("/xl/styles.xml", UriKind.Relative), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
		this.zipArchive.CreateRelationship(wbUri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");

		using var s = e.GetStream();
		using var w = XmlWriter.Create(s);

		w.WriteStartElement("workbook", ns);

		w.WriteStartElement("sheets", ns);
		for (int i = 0; i < this.worksheets.Count; i++)
		{
			var num = i + 1;
			w.WriteStartElement("sheet", ns);

			w.WriteStartAttribute("name");
			w.WriteValue(this.worksheets[i]);
			w.WriteEndAttribute();

			w.WriteStartAttribute("sheetId");
			w.WriteValue(num);
			w.WriteEndAttribute();
			var rel = e.CreateRelationship(new Uri("/xl/worksheets/sheet" + num + ".xml", UriKind.Relative), TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");

			w.WriteStartAttribute("id", RelNS);
			w.WriteValue(rel.Id);
			w.WriteEndAttribute();

			w.WriteEndElement();
		}
		w.WriteEndElement();
		w.WriteEndElement();
	}

	void WriteAppProps()
	{
		var appUri = new Uri("/docProps/app.xml", UriKind.Relative);
		var appEntry = zipArchive.CreatePart(appUri, "application/vnd.openxmlformats-officedocument.extended-properties+xml");
		zipArchive.CreateRelationship(appUri, TargetMode.Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
		using var appStream = appEntry.GetStream();
		using var appX = XmlWriter.Create(appStream);
		appX.WriteStartElement("Properties", PropNS);
		var asmName = Assembly.GetExecutingAssembly().GetName();
		appX.WriteStartElement("Application", PropNS);
		appX.WriteValue(asmName.Name);
		appX.WriteEndElement();
		appX.WriteStartElement("AppVersion", PropNS);
		var v = asmName.Version;
		var ver = v.Major + "." + v.Minor + "." + v.Build;
		appX.WriteValue(ver);
		appX.WriteEndElement();
		appX.WriteEndElement();
	}

	void WriteCoreProps()
	{
		var appUri = new Uri("/docProps/core.xml", UriKind.Relative);
		var appEntry = zipArchive.CreatePart(appUri, "application/vnd.openxmlformats-package.core-properties+xml");
		zipArchive.CreateRelationship(appUri, TargetMode.Internal, "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties");
		using var appStream = appEntry.GetStream();
		using var appX = XmlWriter.Create(appStream);
		appX.WriteStartElement("coreProperties", CoreNS);

		appX.WriteStartElement("lastModifiedBy", CoreNS);
		appX.WriteValue(Environment.UserName);
		appX.WriteEndElement();
		appX.WriteEndElement();
	}

	void WriteStyles()
	{
		var styleUri = new Uri("/xl/styles.xml", UriKind.Relative);
		var appEntry = zipArchive.CreatePart(styleUri, "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
		using var appStream = appEntry.GetStream();
		using var appX = XmlWriter.Create(appStream);
		appX.WriteStartElement("styleSheet", NS);

		appX.WriteStartElement("numFmts", NS);
		for (int i = 0; i < formats.Count; i++)
		{
			appX.WriteStartElement("numFmt", NS);

			appX.WriteStartAttribute("numFmtId");
			appX.WriteValue(fmtOffset + i);
			appX.WriteEndAttribute();

			appX.WriteStartAttribute("formatCode");
			appX.WriteValue(formats[i]);
			appX.WriteEndAttribute();
			appX.WriteEndElement();
		}

		appX.WriteEndElement();


		appX.WriteStartElement("fonts", NS);
		appX.WriteStartElement("font", NS);
		appX.WriteStartElement("name", NS);
		appX.WriteAttributeString("val", "Calibri");
		appX.WriteEndElement();
		appX.WriteEndElement();
		appX.WriteEndElement();

		appX.WriteStartElement("fills");
		appX.WriteStartElement("fill");
		appX.WriteEndElement();
		appX.WriteEndElement();

		appX.WriteStartElement("borders");
		appX.WriteStartElement("border");
		appX.WriteEndElement();
		appX.WriteEndElement();

		appX.WriteStartElement("cellStyleXfs");
		appX.WriteStartElement("xf");
		appX.WriteEndElement();
		appX.WriteEndElement();



		

		appX.WriteStartElement("cellXfs", NS);
		//appX.WriteStartAttribute("count");
		//appX.WriteValue(formats.Count + 1);
		//appX.WriteEndAttribute();

		{
			appX.WriteStartElement("xf", NS);

			appX.WriteStartAttribute("numFmtId");
			appX.WriteValue(0);
			appX.WriteEndAttribute();

			appX.WriteStartAttribute("xfId");
			appX.WriteValue(0);
			appX.WriteEndAttribute();

			appX.WriteEndElement();
		}

		for (int i = 0; i < formats.Count; i++)
		{
			appX.WriteStartElement("xf", NS);

			appX.WriteStartAttribute("numFmtId");
			appX.WriteValue(fmtOffset + i);
			appX.WriteEndAttribute();

			appX.WriteStartAttribute("xfId");
			appX.WriteValue(0);
			appX.WriteEndAttribute();

			appX.WriteEndElement();
		}

		appX.WriteEndElement();


		appX.WriteStartElement("cellStyles");
		appX.WriteStartElement("cellStyle");
		appX.WriteAttributeString("name", "Normal");
		appX.WriteAttributeString("xfId", "0");
		appX.WriteEndElement();
		appX.WriteEndElement();


		appX.WriteEndElement();
	}

	void Close()
	{
		// core.xml isn't needed.
		//WriteCoreProps();
		WriteAppProps();
		WriteSharedStrings();
		WriteStyles();
		WriteWorkbook();
	}

	public override void Dispose()
	{
		this.Close();
		this.zipArchive.Close();
		base.Dispose();
	}
}
