using System.IO;
using System.IO.Compression;
using System.Xml;

namespace Sylvan.Data.Excel;

// supports testing specific xml format scenarios.
class XlsxBuilder
{
	static string BuildSharedStringXml(string[] sharedStrings)
	{
		var sw = new StringWriter();
		using var w = XmlWriter.Create(sw, new XmlWriterSettings { OmitXmlDeclaration = true });
		w.WriteStartElement("sst");
		w.WriteAttributeString("uniqueCount", sharedStrings.Length.ToString());
		foreach (var str in sharedStrings)
		{
			w.WriteStartElement("si");
			w.WriteStartElement("t");
			w.WriteValue(str);
			w.WriteEndElement();
			w.WriteEndElement();
		}
		w.WriteEndElement();
		w.Flush();
		return sw.ToString();
	}

	// creates a minimal .xlsx file that contains the given worksheet xml and sharestrings table.
	public static ExcelDataReader Create(string sheetXml, string sharedStringXml = null, string styleXml = null)
	{
		var ms = new MemoryStream();

		var archive = new ZipArchive(ms, ZipArchiveMode.Create, true);

		if (sharedStringXml != null)
		{
			var ss = archive.CreateEntry("xl/sharedStrings.xml");
			using var os = ss.Open();
			using var tw = new StreamWriter(os);
			tw.Write(sharedStringXml);
		}

		if (styleXml != null)
		{
			var ss = archive.CreateEntry("xl/styles.xml");
			using var os = ss.Open();
			using var tw = new StreamWriter(os);
			tw.Write(styleXml);
		}

		{
			var sheetEntry = archive.CreateEntry("xl/worksheets/sheet1.xml");
			using var os = sheetEntry.Open();
			using var tw = new StreamWriter(os);
			tw.Write(sheetXml);
		}

		{
			var relEntry = archive.CreateEntry("xl/_rels/workbook.xml.rels");
			using var os = relEntry.Open();
			using var w = XmlWriter.Create(os);
			w.WriteStartElement("Relationships");
			w.WriteStartElement("Relationship");
			w.WriteAttributeString("Id", "r1");
			w.WriteAttributeString("Target", "worksheets/sheet1.xml");
			w.WriteEndElement();
			w.WriteEndElement();
		}

		{
			var wbEntry = archive.CreateEntry("xl/workbook.xml");
			using var os = wbEntry.Open();
			using var w = XmlWriter.Create(os);
			w.WriteStartElement("workbook");
			w.WriteStartElement("sheets");
			w.WriteStartElement("sheet");
			w.WriteAttributeString("name", "test sheet");
			w.WriteAttributeString("sheetId", "1");
			w.WriteAttributeString("id", "r1");
			w.WriteEndElement();
			w.WriteEndElement();
			w.WriteEndElement();
		}
		archive.Dispose();
		ms.Seek(0, SeekOrigin.Begin);
		return ExcelDataReader.Create(ms, ExcelWorkbookType.ExcelXml, new ExcelDataReaderOptions { Schema = ExcelSchema.NoHeaders });
	}
}
