using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Xml;

namespace Sylvan.Data.Excel;

// constants and helpers for dealing with open packaging spec
// used for .xlsx and .xlsb
static class OpenPackaging
{
	const string PackageRelationPart = "_rels/.rels";
	const string RelationNS = "http://schemas.openxmlformats.org/package/2006/relationships";

	const string RelationBase = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
	const string DocRelationType = RelationBase + "/officeDocument";
	const string WorksheetRelType = RelationBase + "/worksheet";
	const string StylesRelType = RelationBase + "/styles";
	const string SharedStringsRelType = RelationBase + "/sharedStrings";

	const string StrictRelationBase = "http://purl.oclc.org/ooxml/officeDocument/relationships";
	const string StrictDocRelationType = StrictRelationBase  + "/officeDocument";
	const string StrictWorksheetRelType = StrictRelationBase + "/worksheet";
	const string StrictStylesRelType = StrictRelationBase + "/styles";
	const string StrictSharedStringsRelType = StrictRelationBase + "/sharedStrings";

	const string PropNS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
	internal const string AppPath = "docProps/app.xml";

	internal static readonly XmlWriterSettings XmlSettings =
		new XmlWriterSettings
		{
			OmitXmlDeclaration = true,
			CheckCharacters = false,
		};

	internal static string? GetWorkbookPart(ZipArchive package)
	{
		var part = package.GetEntry(PackageRelationPart);
		if (part == null) return null;

		var doc = new XmlDocument();
		using var stream = part.Open();
		doc.Load(stream);

		var nsm = new XmlNamespaceManager(doc.NameTable);
		nsm.AddNamespace("r", RelationNS);

		var wbPartRel =
			doc.SelectSingleNode($"/r:Relationships/r:Relationship[@Type='{DocRelationType}']", nsm) ??
			doc.SelectSingleNode($"/r:Relationships/r:Relationship[@Type='{StrictDocRelationType}']", nsm);
			
		if (wbPartRel == null) return null;

		var wbPartName = wbPartRel.Attributes?["Target"]?.Value;
		return wbPartName ?? null;
	}

	internal static string GetPartRelationsName(string partName)
	{
		var dir = Path.GetDirectoryName(partName) ?? "";
		//dir = Path.GetRelativePath("/", dir);
		var file = Path.GetFileName(partName);

		return
			string.IsNullOrEmpty(dir)
			? "_rels/" + file + ".rels"
			: dir + "/_rels/" + file + ".rels";
	}

	internal static ZipArchiveEntry? FindEntry(this ZipArchive a, string name)
	{
		var entry = GetEntry(a, name);
		if (entry == null && (name.StartsWith("/") || name.StartsWith("\\")))
		{
			name = name.Substring(1);
			entry = GetEntry(a, name);
		}

		return entry;

		static ZipArchiveEntry? GetEntry(ZipArchive a, string name)
		{
			return a.Entries.FirstOrDefault(e => StringComparer.OrdinalIgnoreCase.Equals(e.FullName, name));
		}
	}

	internal static Dictionary<string, string> LoadWorkbookRelations(ZipArchive package, string workbookPartName, ref string stylesPartName, ref string sharedStringsPartName)
	{
		var workbookPartRelsName = GetPartRelationsName(workbookPartName);

		var part = package.FindEntry(workbookPartRelsName);

		if (part == null) throw new InvalidDataException();

		using Stream sheetRelStream = part.Open();
		var doc = new XmlDocument();
		doc.Load(sheetRelStream);
		if (doc.DocumentElement == null)
		{
			throw new InvalidDataException();
		}
		var nsm = new XmlNamespaceManager(doc.NameTable);
		nsm.AddNamespace("r", doc.DocumentElement.NamespaceURI);
		var nodes = doc.SelectNodes("/r:Relationships/r:Relationship", nsm);

		var root = Path.GetDirectoryName(workbookPartName) ?? "";
		

		static string MakeRelative(string root, string path)
		{
			if (Path.IsPathRooted(path))
			{
				if(path.StartsWith("/") || path.StartsWith("\\"))
				{
					path = path.Substring(1);
				}
				return path;
			}
			return
				root.Length == 0
				? path
				: root + "/" + path;
		}

		var sheetRelMap = new Dictionary<string, string>();
		if (nodes != null)
		{
			foreach (XmlElement node in nodes)
			{
				var id = node.GetAttribute("Id");
				var type = node.GetAttribute("Type");
				var target = node.GetAttribute("Target");
				switch (type)
				{
					case WorksheetRelType:
					case StrictWorksheetRelType:
						var t = MakeRelative(root, target);
						sheetRelMap.Add(id, t);
						break;
					case StylesRelType:
					case StrictStylesRelType:
						stylesPartName = MakeRelative(root, target);
						break;
					case SharedStringsRelType:
					case StrictSharedStringsRelType:
						sharedStringsPartName = MakeRelative(root, target);
						break;
				}
			}
		}
		return sheetRelMap;
	}

	internal static void WriteAppProps(ZipArchive zipArchive)
	{
		var appEntry = zipArchive.CreateEntry(AppPath, CompressionLevel.Fastest);
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
}
