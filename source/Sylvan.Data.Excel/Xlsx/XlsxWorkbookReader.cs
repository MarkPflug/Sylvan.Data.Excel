#nullable enable
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using System.Threading.Tasks;
using System.Threading;
using Sylvan.Data.Excel.Xlsx;

#if !SPAN
using ReadonlyCharSpan = System.String;
#else
using ReadonlyCharSpan = System.ReadOnlySpan<char>;
using CharSpan = System.Span<char>;
#endif

namespace Sylvan.Data.Excel;

sealed class XlsxWorkbookReader : ExcelDataReader
{
	readonly ZipArchive package;

	Stream sheetStream;
	XmlReader? reader;

	StringBuilder? stringBuilder;

	bool hasRows;
	//bool skipEmptyRows = true; // TODO: make this an option?

	char[] valueBuffer = new char[64];

	int rowIndex;
	int parsedRowIndex = -1;
	// the number of fields in the parsedRowIndex.
	int curFieldCount = -1;

	public override ExcelWorkbookType WorkbookType => ExcelWorkbookType.ExcelXml;

	const string DefaultWorkbookPartName = "xl/workbook.xml";

	public XlsxWorkbookReader(Stream iStream, ExcelDataReaderOptions opts) : base(iStream, opts)
	{
		this.rowCount = -1;
		this.sheetStream = Stream.Null;

		package = new ZipArchive(iStream, ZipArchiveMode.Read, true);

		var workbookPartName = OpenPackaging.GetWorkbookPart(package) ?? DefaultWorkbookPartName;

		var workbookPart = package.FindEntry(workbookPartName);

		if (workbookPart == null)
			throw new InvalidDataException();

		var stylesPartName = "xl/styles.xml";
		var sharedStringsPartName = "xl/sharedStrings.xml";

		var sheetRelMap = OpenPackaging.LoadWorkbookRelations(package, workbookPartName, ref stylesPartName, ref sharedStringsPartName);

		var ssPart = package.FindEntry(sharedStringsPartName);
		var stylePart = package.FindEntry(stylesPartName);

		LoadSharedStrings(ssPart);

		using (Stream sheetsStream = workbookPart.Open())
		{
			// quick and dirty, good enough, this doc should be small.
			var doc = new XmlDocument();
			doc.Load(sheetsStream);

			if (doc.DocumentElement == null)
			{
				throw new InvalidDataException();
			}

			var nsm = new XmlNamespaceManager(doc.NameTable);
			var ns = doc.DocumentElement.NamespaceURI;
			nsm.AddNamespace("x", ns);
			var nodes = doc.SelectNodes("/x:workbook/x:sheets/x:sheet", nsm);
			var prNode = doc.SelectSingleNode("/x:workbook/x:workbookPr", nsm);
			if (prNode?.Attributes?["date1904"]?.Value == "1")
			{
				this.dateMode = DateMode.Mode1904;
			}

			var sheets = new List<SheetInfo>();
			if (nodes != null)
			{
				foreach (XmlElement sheetElem in nodes)
				{
					var id = int.Parse(sheetElem.GetAttribute("sheetId"));
					var name = sheetElem.GetAttribute("name");
					var state = sheetElem.GetAttribute("state");
					var refId = sheetElem.Attributes.OfType<XmlAttribute>().Single(a => a.LocalName == "id").Value;

					var hidden = StringComparer.OrdinalIgnoreCase.Equals(state, "hidden");
					var si = new SheetInfo(name, sheetRelMap[refId], hidden);
					sheets.Add(si);
				}
			}
			this.sheetInfos = sheets.ToArray();
		}
		if (stylePart == null)
		{

			this.xfMap = Array.Empty<int>();
			this.formats = ExcelFormat.CreateFormatCollection();
		}
		else
		{
			using (Stream styleStream = stylePart.Open())
			{
				var doc = new XmlDocument();
				doc.Load(styleStream);
				if (doc.DocumentElement == null)
				{
					throw new InvalidDataException();
				}
				var nsm = new XmlNamespaceManager(doc.NameTable);
				var ns = doc.DocumentElement.NamespaceURI;
				nsm.AddNamespace("x", ns);
				var nodes = doc.SelectNodes("/x:styleSheet/x:numFmts/x:numFmt", nsm);
				this.formats = ExcelFormat.CreateFormatCollection();
				if (nodes != null)
				{
					foreach (XmlElement fmt in nodes)
					{
						var id = int.Parse(fmt.GetAttribute("numFmtId"));
						var str = fmt.GetAttribute("formatCode");
						var ef = new ExcelFormat(str);
						if (formats.ContainsKey(id))
						{

						}
						else
						{
							formats[id] = ef;
						}
					}
				}
				XmlElement? xfsElem = (XmlElement?)doc.SelectSingleNode("/x:styleSheet/x:cellXfs", nsm);
				if (xfsElem != null)
				{
					this.xfMap = new int[xfsElem.ChildNodes.Count];
					for (int idx = 0; idx < xfMap.Length; idx++)
					{
						var xf = (XmlElement?)xfsElem.ChildNodes[idx];
						if (xf == null)
						{
							throw new InvalidDataException();
						}
						var fmtIdStr = xf.GetAttribute("numFmtId");
						xfMap[idx] = int.Parse(fmtIdStr);
					}
				}
				else
				{
					this.xfMap = Array.Empty<int>();
				}
			}
		}

		NextResult();
	}

	private protected override ref readonly FieldInfo GetFieldValue(int ordinal)
	{
		if (rowIndex < this.parsedRowIndex)
		{
			return ref FieldInfo.Null;
		}
		return ref base.GetFieldValue(ordinal);
	}

	private protected override Task<bool> OpenWorksheetAsync(int sheetIdx, CancellationToken cancel)
	{
		var sheetName = sheetInfos[sheetIdx].Part;
		// the relationship is recorded as an absolute path
		// but the zip entry has a relative name.
		sheetName = sheetName.TrimStart('/');
		var sheetPart = package.GetEntry(sheetName);
		if (sheetPart == null)
			return Task.FromResult(false);

		this.sheetStream = sheetPart.Open();

		var tr = new StreamReader(this.sheetStream, Encoding.UTF8, true, 0x10000);

		var settings = new XmlReaderSettings
		{
			CheckCharacters = false,
			ValidationType = ValidationType.None,
			ValidationFlags = System.Xml.Schema.XmlSchemaValidationFlags.None,
#if SPAN
			NameTable = new SheetNameTable(),
#endif
		};

		this.reader = XmlReader.Create(tr, settings);

		// worksheet
		while (reader.Read())
		{
			if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "worksheet")
			{
				break;
			}
		}

		while (reader.Read())
		{
			if (reader.NodeType == XmlNodeType.Element)
			{
				if (reader.LocalName == "dimension")
				{
					var dim = reader.GetAttribute("ref");
					if (dim != null)
					{
						var idx = dim.IndexOf(':');
						if (CellPosition.TryParse(dim.Substring(idx + 1), out var p))
						{
							this.rowCount = p.Row + 1;
						}
					}
				}
				if (reader.LocalName == "sheetData")
				{
					break;
				}
			}
		}

		this.hasRows = InitializeSheet();
		this.sheetIdx = sheetIdx;
		return Task.FromResult(true);
	}

	public override bool NextResult()
	{
		sheetIdx++;
		for (; sheetIdx < this.sheetInfos.Length; sheetIdx++)
		{
			if (readHiddenSheets || sheetInfos[sheetIdx].Hidden == false)
			{
				break;
			}
		}
		if (sheetIdx >= this.sheetInfos.Length)
		{
			return false;
		}

		return OpenWorksheetAsync(sheetIdx, default).GetAwaiter().GetResult();
	}

	bool InitializeSheet()
	{
		this.state = State.Initializing;

		if (reader == null)
		{
			this.state = State.Closed;
			throw new InvalidOperationException();
		}
		this.parsedRowIndex = -1;
		if (!NextRow())
		{
			return false;
		}
		ParseRowValues();

		if (this.parsedRowIndex > 0)
		{
			this.curFieldCount = rowFieldCount;
			this.rowFieldCount = 0;
		}

		if (LoadSchema())
		{
			this.state = State.Initialized;
			this.rowIndex = -1;
		}
		else
		{
			this.state = State.Open;
			this.rowIndex = 0;
		}

		return true;
	}

	bool NextRow()
	{
		var ci = NumberFormatInfo.InvariantInfo;
		if (ReadToFollowing(reader!, "row"))
		{
			if (reader!.MoveToAttribute("r"))
			{
				int row;
#if SPAN
				var len = reader.ReadValueChunk(valueBuffer, 0, valueBuffer.Length);
				if (len < valueBuffer.Length && int.TryParse(valueBuffer.AsSpan(0, len), NumberStyles.Integer, ci, out row))
				{
				}
				else
				{
					throw new FormatException();
				}
#else
				var str = reader.Value;
				row = int.Parse(str, NumberStyles.Integer, ci);
#endif

				this.parsedRowIndex = row - 1;
			}
			else
			{
				this.parsedRowIndex = -1;
			}
			reader.MoveToElement();
			return true;
		}
		return false;
	}

	struct CellPosition
	{
		public int Column;
		public int Row;

		public static bool TryParse(ReadonlyCharSpan str, out CellPosition pos)
		{
			pos = default;

			int col = -1;
			int row = 0;
			int i = 0;
			char c;
			for (; i < str.Length; i++)
			{
				c = str[i];
				var v = c - 'A';
				if ((uint)v < 26u)
				{
					col = ((col + 1) * 26) + v;
				}
				else
				{
					break;
				}
			}

			for (; i < str.Length; i++)
			{
				c = str[i];
				var v = c - '0';
				if ((uint)v >= 10)
				{
					return false;
				}
				row = row * 10 + v;
			}
			pos = new CellPosition() { Column = col, Row = row - 1 };
			return true;
		}
	}

	public override bool Read()
	{
		rowIndex++;
		if (state == State.Open)
		{
			if (rowIndex <= parsedRowIndex)
			{
				if (rowIndex < parsedRowIndex)
				{
					this.rowFieldCount = 0;
				} 
				else
				{
					this.rowFieldCount = curFieldCount;
					this.curFieldCount = -1;
				}
				return true;
			}
			while (NextRow())
			{
				var c = ParseRowValues();
				if (c == 0)
				{
					// handles trailing empty rows.
					continue;
				}
				if (rowIndex < parsedRowIndex)
				{
					this.curFieldCount = c;
					this.rowFieldCount = 0;
					return true;
				}				
				return true;
			}
		}
		else
		if (state == State.Initialized)
		{
			// after initizialization, the first record would already be in the buffer
			// if hasRows is true.
			if (hasRows)
			{
				this.state = State.Open;
				if (curFieldCount >= 0)
				{
					this.rowFieldCount = curFieldCount;
					this.curFieldCount = -1;
				}
				return true;
			}
		}
		rowIndex = -1;
		this.state = State.End;
		return false;
	}

	int ParseRowValues()
	{
		var ci = NumberFormatInfo.InvariantInfo;
		XmlReader reader = this.reader!;
		FieldInfo[] values = this.values;
		int len;

		Array.Clear(this.values, 0, this.values.Length);
		if (!ReadToDescendant(reader, "c"))
		{
			return 0;
		}

		int valueCount = 0;
		var col = -1;

		do
		{
			CellType type = CellType.Numeric;
			int xfIdx = 0;
			col++;
			while (reader.MoveToNextAttribute())
			{
				var n = reader.Name;
				if (n == "r")
				{
					len = reader.ReadValueChunk(valueBuffer, 0, valueBuffer.Length);
					if (CellPosition.TryParse(valueBuffer.AsSpan().ToParsable(0, len), out var pos))
					{
						col = pos.Column;
					}
				}
				else
				if (n == "t")
				{
					len = reader.ReadValueChunk(valueBuffer, 0, valueBuffer.Length);
					type = GetCellType(valueBuffer, len);
				}
				else
				if (n == "s")
				{
					len = reader.ReadValueChunk(valueBuffer, 0, valueBuffer.Length);
					if (!TryParse(valueBuffer.AsSpan().ToParsable(0, len), out xfIdx))
					{
						throw new FormatException();
					}
				}
			}

			if (col >= values.Length)
			{
				Array.Resize(ref values, col + 8);
				this.values = values;
			}

			static bool TryParse(ReadonlyCharSpan span, out int value)
			{
				int a = 0;
				for (int i = 0; i < span.Length; i++)
				{
					var d = span[i] - '0';
					if ((uint)d >= 10)
					{
						value = 0;
						return false;
					}
					a = a * 10 + d;
				}
				value = a;
				return true;
			}

			static CellType GetCellType(char[] b, int l)
			{
				switch (b[0])
				{
					case 'b':
						return CellType.Boolean;
					case 'e':
						return CellType.Error;
					case 's':
						return l == 1 ? CellType.SharedString : CellType.String;
					case 'i':
						return CellType.InlineString;
					case 'd':
						return CellType.Date;
					case 'n':
						return CellType.Numeric;
					default:
						// TODO:
						throw new InvalidDataException();
				}
			}

			ref FieldInfo fi = ref values[col];
			fi.xfIdx = xfIdx;

			reader.MoveToElement();
			var depth = reader.Depth;

			if (type == CellType.InlineString)
			{
				if (ReadToDescendant(reader, "is"))
				{
					fi.strValue = ReadString(reader);
					fi.type = ExcelDataType.String;
					valueCount++;
					this.rowFieldCount = col + 1;
				}
				else
				{
					fi.strValue = string.Empty;
					fi.type = ExcelDataType.Null;
				}
			}
			else
			if (ReadToDescendant(reader, "v"))
			{

				if (!reader.IsEmptyElement)
				{
					reader.Read();
				}
				switch (type)
				{
					case CellType.Numeric:
						fi.type = ExcelDataType.Numeric;
#if SPAN
						len = reader.ReadValueChunk(valueBuffer, 0, valueBuffer.Length);
						if (len < valueBuffer.Length && double.TryParse(valueBuffer.AsSpan(0, len), NumberStyles.Float, ci, out fi.numValue))
						{
						}
						else
						{
							throw new FormatException();
						}
#else
						var str = reader.Value;
						fi.numValue = double.Parse(str, NumberStyles.Float, ci);
#endif
						break;
					case CellType.Date:
						len = reader.ReadValueChunk(valueBuffer, 0, valueBuffer.Length);
						if (len < valueBuffer.Length)
						{
							if (!IsoDate.TryParse(valueBuffer.AsSpan(0, len), out fi.dtValue))
							{
								throw new FormatException();
							}
						}
						else
						{
							throw new FormatException();
						}
						fi.type = ExcelDataType.DateTime;
						break;
					case CellType.SharedString:
						if (reader.NodeType == XmlNodeType.Text)
						{
							len = reader.ReadValueChunk(valueBuffer, 0, valueBuffer.Length);
							if (len >= valueBuffer.Length)
								throw new FormatException();
							if (!TryParse(valueBuffer.AsSpan().ToParsable(0, len), out int strIdx))
							{
								throw new FormatException();
							}
							fi.strValue = GetSharedString(strIdx);
						}
						else
						{
							fi.strValue = string.Empty;
						}
						fi.type = fi.strValue.Length == 0 ? ExcelDataType.Null : ExcelDataType.String;
						break;
					case CellType.String:
						if (reader.NodeType == XmlNodeType.Text)
						{
							var s = reader.ReadContentAsString();
							if (reader.XmlSpace != XmlSpace.Preserve)
							{
								s = s.Trim();
							}
							fi.strValue = s;
							fi.type = ExcelDataType.String;
						}
						else
						{
							fi.strValue = string.Empty;
							fi.type = ExcelDataType.Null;
						}
						break;
					case CellType.InlineString:
						fi.strValue = ReadString(reader);
						fi.type = ExcelDataType.String;
						if (fi.strValue.Length == 0)
						{
							fi.type = ExcelDataType.Null;
						}
						break;
					case CellType.Boolean:
						len = reader.ReadValueChunk(valueBuffer, 0, valueBuffer.Length);
						if (len < 1)
						{
							throw new FormatException();
						}
						fi.type = ExcelDataType.Boolean;
						fi = new FieldInfo(valueBuffer[0] != '0');
						break;
					case CellType.Error:
						len = reader.ReadValueChunk(valueBuffer, 0, valueBuffer.Length);
						fi = new FieldInfo(GetErrorCode(valueBuffer.AsSpan(0, len)));
						break;
					default:
						throw new InvalidDataException();
				}
				if (fi.type != ExcelDataType.Null)
				{
					valueCount++;
					this.rowFieldCount = col + 1;
				}
			}

			while (reader.Depth > depth)
			{
				reader.Read();
			}

		} while (ReadToNextSibling(reader, "c"));
		return valueCount == 0 ? 0 : col + 1;
	}

	static bool Equal(CharSpan l, ReadonlyCharSpan r)
	{
		if (l.Length != r.Length) return false;
		for (int i = 0; i < l.Length; i++)
		{
			if (l[i] != r[i])
				return false;
		}
		return true;
	}

	static ExcelErrorCode GetErrorCode(CharSpan str)
	{
		if (Equal(str, "#DIV/0!"))
			return ExcelErrorCode.DivideByZero;
		if (Equal(str, "#VALUE!"))
			return ExcelErrorCode.Value;
		if (Equal(str, "#REF!"))
			return ExcelErrorCode.Reference;
		if (Equal(str, "#NAME?"))
			return ExcelErrorCode.Name;
		if (Equal(str, "#N/A"))
			return ExcelErrorCode.NotAvailable;
		if (Equal(str, "#NULL!"))
			return ExcelErrorCode.Null;

		throw new FormatException();
	}

	static bool ReadToFollowing(XmlReader reader, string localName)
	{
		while (reader.Read())
		{
			if (reader.NodeType == XmlNodeType.Element && localName == reader.LocalName)
			{
				return true;
			}
		}
		return false;
	}

	static bool ReadToNextSibling(XmlReader reader, string localName)
	{
		while (SkipSubtree(reader))
		{
			XmlNodeType nodeType = reader.NodeType;
			if (nodeType == XmlNodeType.Element && localName == reader.LocalName)
			{
				return true;
			}
			if (nodeType == XmlNodeType.EndElement || reader.EOF)
			{
				break;
			}
		}
		return false;
	}

	static bool ReadToDescendant(XmlReader reader, string localName)
	{
		int num = reader.Depth;
		if (reader.NodeType != XmlNodeType.Element)
		{
			if (reader.ReadState != 0)
			{
				return false;
			}
			num--;
		}
		else if (reader.IsEmptyElement)
		{
			return false;
		}
		while (reader.Read() && reader.Depth > num)
		{
			if (reader.NodeType == XmlNodeType.Element && localName == reader.LocalName)
			{
				return true;
			}
		}
		return false;
	}

	static bool SkipSubtree(XmlReader reader)
	{
		reader.MoveToElement();
		if (reader.NodeType == XmlNodeType.Element && !reader.IsEmptyElement)
		{
			int depth = reader.Depth;
			while (reader.Read() && depth < reader.Depth)
			{
			}
			if (reader.NodeType == XmlNodeType.EndElement)
			{
				return reader.Read();
			}
			return false;
		}
		return reader.Read();
	}

	enum CellType
	{
		Numeric,
		String,
		SharedString,
		InlineString,
		Boolean,
		Error,
		Date,
	}

	internal override DateTime GetDateTimeValue(int ordinal)
	{
		return this.values[ordinal].dtValue;
	}

	public override int MaxFieldCount => 16384;

	public override int RowNumber => rowIndex + 1;

	void LoadSharedStrings(ZipArchiveEntry? entry)
	{
		if (entry == null)
		{
			return;
		}
		using Stream ssStream = entry.Open();

		var settings = new XmlReaderSettings
		{
			CheckCharacters = false,
#if SPAN
			// name table optimization requires ROS
			NameTable = new SharedStringsNameTable(),
#endif
		};

		using var reader = XmlReader.Create(ssStream, settings);

		while (reader.Read())
		{
			if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "sst")
			{
				break;
			}
		}

		var countStr = reader.GetAttribute("uniqueCount");

		var count = 0;
		if (!string.IsNullOrEmpty(countStr) && int.TryParse(countStr, out count) && count >= 0)
		{

		}
		else
		{
			// try to estimate the number of strings based on the entry size
			// Estimate ~24 bytes per string record.
			var estimatedCount = (int)(entry.Length / 24);
			count = Math.Max(1, estimatedCount);
		}

		var sstList = new List<string>(count);

		while (reader.Read())
		{
			if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "si")
			{
				var str = ReadString(reader);
				sstList.Add(str);
			}
		}

		this.sst = sstList.ToArray();
	}

	string ReadString(XmlReader reader)
	{
		this.stringBuilder ??= new StringBuilder();

		var empty = reader.IsEmptyElement;
		string str = string.Empty;
		if (empty)
		{
			reader.ReadEndElement();
		}
		else
		{
			var depth = reader.Depth;
			int c = 0;
			while (reader.Read() && reader.Depth > depth)
			{
				if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "t")
				{
					string? s = string.Empty;
					var wsPreserve = reader.XmlSpace == XmlSpace.Preserve;
					if (reader.IsEmptyElement)
					{
						s = string.Empty;
					}
					else
					{
						reader.Read();
						if (reader.NodeType == XmlNodeType.Text || reader.NodeType == XmlNodeType.SignificantWhitespace || reader.NodeType == XmlNodeType.Whitespace)
						{
							var val = reader.Value;
							if (!wsPreserve)
							{
								val = val.Trim();
							}
							s = OpenXmlCodec.DecodeString(val);
						}
						else if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "t")
						{

						}
						else
						{
							throw new InvalidDataException();
						}
					}
					if (c == 0)
					{
						str = s;
					}
					else
					if (c == 1)
					{
						stringBuilder.Clear();
						stringBuilder.Append(str);
						stringBuilder.Append(s);
					}
					else
					{
						stringBuilder.Append(s);
					}
					c++;
				}
			}
			if (c > 1)
			{
				str = stringBuilder.ToString();
			}
		}
		return str;
	}

	string GetSharedString(int i)
	{
		if ((uint)i >= sst.Length)
			throw new ArgumentOutOfRangeException(nameof(i));

		return sst[i];
	}
}
