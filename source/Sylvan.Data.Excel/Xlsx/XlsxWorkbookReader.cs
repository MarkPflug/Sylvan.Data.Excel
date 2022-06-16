#nullable enable
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;

#if !NETSTANDARD2_1_OR_GREATER
using ReadonlyCharSpan = System.String;
using CharSpan = System.Text.StringBuilder;
#else
using ReadonlyCharSpan = System.ReadOnlySpan<char>;
using CharSpan = System.Span<char>;
#endif

namespace Sylvan.Data.Excel;

sealed class XlsxWorkbookReader : ExcelDataReader
{
	readonly ZipArchive package;
	Dictionary<int, ExcelFormat> formats;
	int[] xfMap;
	int sheetIdx = -1;
	int rowCount;

	bool readHiddenSheets;

	Stream sheetStream;
	XmlReader? reader;

	FieldInfo[] values;
	int rowFieldCount;
	State state;
	bool hasRows;
	bool skipEmptyRows = true; // TODO: make this an option?
	SheetInfo[] sheetNames;
	bool[] sheetHiddenFlags;
	bool errorAsNull;

	string refName;
	string typeName;
	string styleName;
	string rowName;
	string valueName;
	string inlineStringName;
	string cellName;
	string sheetNS;

	char[] valueBuffer = new char[64];

	int rowIndex;
	int parsedRowIndex = -1;

	struct FieldInfo
	{
		public ExcelDataType type;
		public string strValue;
		public double numValue;
		public DateTime dtValue;
		public int xfIdx;
	}

	public override int RowCount => rowCount;

	public override ExcelWorkbookType WorkbookType => ExcelWorkbookType.ExcelXml;

	class SheetInfo
	{
		public SheetInfo(string name, string part)
		{
			this.Name = name;
			this.Part = part;
		}

		public string Name { get; }
		public string Part { get; }
	}

	static ZipArchiveEntry? GetEntry(ZipArchive a, string name)
	{
		return a.Entries.FirstOrDefault(e => StringComparer.OrdinalIgnoreCase.Equals(e.FullName, name));
	}

	public XlsxWorkbookReader(Stream iStream, ExcelDataReaderOptions opts) : base(iStream, opts.Schema)
	{
		this.rowCount = -1;
		this.values = Array.Empty<FieldInfo>();
		this.sheetStream = Stream.Null;

		this.rowName = this.cellName = this.valueName = this.refName = this.styleName = this.typeName = this.inlineStringName = string.Empty;
		this.sheetNS = string.Empty;
		this.errorAsNull = opts.GetErrorAsNull;
		this.readHiddenSheets = opts.ReadHiddenWorksheets;

		package = new ZipArchive(iStream, ZipArchiveMode.Read);

		var ssPart = GetEntry(package, "xl/sharedStrings.xml");
		var stylePart = GetEntry(package, "xl/styles.xml");

		var sheetsPart = GetEntry(package, "xl/workbook.xml");
		var sheetsRelsPart = GetEntry(package, "xl/_rels/workbook.xml.rels");
		if (sheetsPart == null || sheetsRelsPart == null)
			throw new InvalidDataException();

		this.stringData = Array.Empty<string>();
		LoadSharedStrings(ssPart);

		var sheetNameList = new List<SheetInfo>();
		var sheetHiddenList = new List<bool>();
		Dictionary<string, string> sheetRelMap = new Dictionary<string, string>();
		using (Stream sheetRelStream = sheetsRelsPart.Open())
		{
			var doc = new XmlDocument();
			doc.Load(sheetRelStream);
			var nsm = new XmlNamespaceManager(doc.NameTable);
			nsm.AddNamespace("r", doc.DocumentElement.NamespaceURI);
			var nodes = doc.SelectNodes("/r:Relationships/r:Relationship", nsm);
			foreach (XmlElement node in nodes)
			{
				var id = node.GetAttribute("Id");
				var target = node.GetAttribute("Target");
				if (target.StartsWith("/"))
				{

				}
				else
				{
					target = "xl/" + target;
				}

				sheetRelMap.Add(id, target);
			}
		}

		using (Stream sheetsStream = sheetsPart.Open())
		{
			// quick and dirty, good enough, this doc should be small.
			var doc = new XmlDocument();
			doc.Load(sheetsStream);
			var nsm = new XmlNamespaceManager(doc.NameTable);
			var ns = doc.DocumentElement.NamespaceURI;
			nsm.AddNamespace("x", ns);
			var nodes = doc.SelectNodes("/x:workbook/x:sheets/x:sheet", nsm);
			foreach (XmlElement sheetElem in nodes)
			{
				var id = int.Parse(sheetElem.GetAttribute("sheetId"));
				var name = sheetElem.GetAttribute("name");
				var state = sheetElem.GetAttribute("state");
				var refId = sheetElem.Attributes.OfType<XmlAttribute>().Single(a => a.LocalName == "id").Value;

				sheetHiddenList.Add(StringComparer.OrdinalIgnoreCase.Equals(state, "hidden"));
				var si = new SheetInfo(name, sheetRelMap[refId]);
				sheetNameList.Add(si);
			}
		}
		this.sheetNames = sheetNameList.ToArray();
		this.sheetHiddenFlags = sheetHiddenList.ToArray();
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
				var nsm = new XmlNamespaceManager(doc.NameTable);
				var ns = doc.DocumentElement.NamespaceURI;
				nsm.AddNamespace("x", ns);
				var nodes = doc.SelectNodes("/x:styleSheet/x:numFmts/x:numFmt", nsm);
				this.formats = ExcelFormat.CreateFormatCollection();
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

				XmlElement xfsElem = (XmlElement)doc.SelectSingleNode("/x:styleSheet/x:cellXfs", nsm);
				if (xfsElem != null)
				{
					var c =
						xfsElem.HasAttribute("count")
						? int.Parse(xfsElem.GetAttribute("count"))
						: 0;

					this.xfMap = new int[c];
					int idx = 0;
					foreach (XmlElement xf in xfsElem.ChildNodes)
					{
						var id = int.Parse(xf.GetAttribute("numFmtId"));
						var apply = xf.HasAttribute("applyNumberFormat");
						xfMap[idx] = apply ? id : 0;
						idx++;
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

	XmlReaderSettings settings = new XmlReaderSettings()
	{
		CheckCharacters = false,
		ValidationType = ValidationType.None,
		ValidationFlags = System.Xml.Schema.XmlSchemaValidationFlags.None,
	};

	public override bool NextResult()
	{
		sheetIdx++;
		for (; sheetIdx < this.sheetNames.Length; sheetIdx++)
		{
			if (readHiddenSheets || sheetHiddenFlags[sheetIdx] == false)
			{
				break;
			}
		}
		if (sheetIdx >= this.sheetNames.Length)
		{
			return false;
		}
		var sheetName = sheetNames[sheetIdx].Part;
		// the relationship is recorded as an absolute path
		// but the zip entry has a relative name.
		sheetName = sheetName.TrimStart('/');
		var sheetPart = package.GetEntry(sheetName);
		if (sheetPart == null)
			return false;

		this.sheetStream = sheetPart.Open();

		var tr = new StreamReader(this.sheetStream, Encoding.UTF8, true, 0x10000);

		this.reader = XmlReader.Create(tr, settings);
		refName = this.reader.NameTable.Add("r");
		typeName = this.reader.NameTable.Add("t");
		styleName = this.reader.NameTable.Add("s");
		rowName = this.reader.NameTable.Add("row");
		valueName = this.reader.NameTable.Add("v");
		inlineStringName = this.reader.NameTable.Add("is");
		cellName = this.reader.NameTable.Add("c");

		// worksheet
		while (reader.Read())
		{
			if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "worksheet")
			{
				sheetNS = this.reader.NameTable.Add(reader.NamespaceURI);
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
					var idx = dim.IndexOf(':');
					var p = CellPosition.Parse(dim.Substring(idx + 1));
					this.rowCount = p.Row + 1;
				}
				if (reader.LocalName == "sheetData")
				{
					break;
				}
			}
		}

		this.hasRows = InitializeSheet();
		return true;
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
		this.rowIndex = 0;

		var hasHeaders = schema.HasHeaders(this.WorksheetName!);

		LoadSchema(!hasHeaders);

		if (hasHeaders)
		{
			this.state = State.Open;
			Read();
			this.rowIndex = 0;
		}
		else
		{
			this.rowIndex = -1;
		}

		this.state = State.Initialized;
		return true;
	}

	bool NextRow()
	{
		var ci = NumberFormatInfo.InvariantInfo;
		if (ReadToFollowing(reader!, rowName))
		{
			if (reader!.MoveToAttribute(refName))
			{
				int row;
#if SPAN_PARSE
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
				this.parsedRowIndex++;
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

		public static CellPosition Parse(ReadonlyCharSpan str)
		{
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
					throw new FormatException();
				}
				row = row * 10 + v;
			}
			return new CellPosition() { Column = col, Row = row - 1 };
		}
	}

	public override bool Read()
	{
	start:
		rowIndex++;
		if (state == State.Open)
		{
			if (rowIndex <= parsedRowIndex)
				return true;
			while (NextRow())
			{
				var c = ParseRowValues();
				if (c == 0 && skipEmptyRows)
				{
					continue;
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
				if (this.RowFieldCount == 0 && skipEmptyRows)
					goto start;
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
		if (!ReadToDescendant(reader, cellName))
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
				if (ReferenceEquals(n, refName))
				{
					len = reader.ReadValueChunk(valueBuffer, 0, valueBuffer.Length);
					var pos = CellPosition.Parse(valueBuffer.AsSpan(0, len));
					if (pos.Column >= 0)
					{
						col = pos.Column;
					}
				}
				else
				if (ReferenceEquals(n, typeName))
				{
					len = reader.ReadValueChunk(valueBuffer, 0, valueBuffer.Length);
					type = GetCellType(valueBuffer, len);
				}
				else
				if (ReferenceEquals(n, styleName))
				{
					len = reader.ReadValueChunk(valueBuffer, 0, valueBuffer.Length);
					if (!TryParse(valueBuffer.AsSpan(0, len), out xfIdx))
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
				valueCount++;
				this.rowFieldCount = col + 1;
				if (ReadToDescendant(reader, inlineStringName))
				{
					fi.strValue = ReadString(reader);
					fi.type = ExcelDataType.String;
				} 
				else
				{
					throw new InvalidDataException();
				}
			}
			else
			if (ReadToDescendant(reader, valueName))
			{
				valueCount++;
				this.rowFieldCount = col + 1;
				reader.Read();
				switch (type)
				{
					case CellType.Numeric:
						fi.type = ExcelDataType.Numeric;
#if SPAN_PARSE
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
						len = reader.ReadValueChunk(valueBuffer, 0, valueBuffer.Length);
						if (len >= valueBuffer.Length)
							throw new FormatException();
						if (!TryParse(valueBuffer.AsSpan(0, len), out int strIdx))
						{
							throw new FormatException();
						}
						fi.strValue = GetSharedString(strIdx);
						fi.type = ExcelDataType.String;
						break;
					case CellType.String:
						fi.strValue = reader.ReadContentAsString();
						fi.type = ExcelDataType.String;
						break;
					case CellType.InlineString:
						fi.strValue = ReadString(reader);
						fi.type = ExcelDataType.String;
						break;
					case CellType.Boolean:
						fi.strValue = reader.ReadContentAsString();
						fi.type = ExcelDataType.Boolean;
						break;
					case CellType.Error:
						fi.strValue = reader.ReadContentAsString();
						fi.type = ExcelDataType.Error;
						break;
					default:
						throw new InvalidDataException();
				}
			}

			while (reader.Depth > depth)
			{
				reader.Read();
			}

		} while (ReadToNextSibling(reader, cellName));
		return valueCount == 0 ? 0 : col + 1;
	}

	static bool ReadToFollowing(XmlReader reader, string localName)
	{
		while (reader.Read())
		{
			if (reader.NodeType == XmlNodeType.Element && object.ReferenceEquals(localName, reader.LocalName))
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
			if (nodeType == XmlNodeType.Element && object.ReferenceEquals(localName, reader.LocalName))
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
			if (reader.NodeType == XmlNodeType.Element && object.ReferenceEquals(localName, reader.LocalName))
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

	public override int GetOrdinal(string name)
	{
		if (this.columnSchema == null)
			return -1;
		for (int i = 0; i < this.columnSchema.Count; i++)
		{
			if (this.columnSchema[i].ColumnName == name)
				return i;
		}
		return -1;
	}

	public override ExcelDataType GetExcelDataType(int ordinal)
	{
		AssertRange(ordinal);
		if (rowIndex < parsedRowIndex)
			return ExcelDataType.Null;
		if (ordinal >= this.rowFieldCount)
			return ExcelDataType.Null;
		return values[ordinal].type;
	}

	static ExcelErrorCode GetErrorCode(string str)
	{
		if (str == "#DIV/0!")
			return ExcelErrorCode.DivideByZero;
		if (str == "#VALUE!")
			return ExcelErrorCode.Value;
		if (str == "#REF!")
			return ExcelErrorCode.Reference;
		if (str == "#NAME?")
			return ExcelErrorCode.Name;
		if (str == "#N/A")
			return ExcelErrorCode.NotAvailable;
		if (str == "#NULL!")
			return ExcelErrorCode.Null;

		throw new FormatException();
	}

	public override bool GetBoolean(int ordinal)
	{
		var fi = this.values[ordinal];
		switch (fi.type)
		{
			case ExcelDataType.Boolean:
				return fi.strValue[0] == '1';
			case ExcelDataType.Numeric:
				return this.GetDouble(ordinal) != 0;
			case ExcelDataType.String:
				return bool.TryParse(fi.strValue, out var b)
					? b
					: throw new FormatException();
			case ExcelDataType.Error:
				var code = GetErrorCode(fi.strValue);
				throw new ExcelFormulaException(ordinal, RowNumber, code);
		}
		throw new InvalidCastException();
	}

	internal override DateTime GetDateTimeValue(int ordinal)
	{
		return this.values[ordinal].dtValue;
	}

	public override double GetDouble(int ordinal)
	{
		if (rowIndex == parsedRowIndex)
		{
			ref var fi = ref values[ordinal];
			var type = fi.type;
			switch (type)
			{
				case ExcelDataType.Numeric:
					return fi.numValue;
				case ExcelDataType.String:
					return double.Parse(fi.strValue);
				case ExcelDataType.Error:
					throw Error(ordinal);
			}
		}
		throw new InvalidCastException();
	}

	ExcelFormulaException Error(int ordinal)
	{
		return new ExcelFormulaException(ordinal, rowIndex, GetFormulaError(ordinal));
	}

	public override string GetString(int ordinal)
	{
		AssertRange(ordinal);
		if (rowIndex < parsedRowIndex)
		{
			return string.Empty;
		}
		if (ordinal > this.rowFieldCount)
		{
			return string.Empty;
		}

		ref var fi = ref values[ordinal];
		switch (fi.type)
		{
			case ExcelDataType.Error:
				if (this.errorAsNull)
				{
					return string.Empty;
				}
				throw Error(ordinal);
			case ExcelDataType.Boolean:
				return fi.strValue[0] == '0' ? bool.FalseString : bool.TrueString;
			case ExcelDataType.Numeric:
				return FormatVal(fi.xfIdx, fi.numValue);
			case ExcelDataType.DateTime:
				return IsoDate.ToStringIso(fi.dtValue);
		}
		return fi.strValue ?? string.Empty;
	}

	string FormatVal(int xfIdx, double val)
	{
		var fmtIdx = xfIdx >= this.xfMap.Length ? -1 : this.xfMap[xfIdx];
		if (fmtIdx == -1)
		{
			return val.ToString();
		}

		if (formats.TryGetValue(fmtIdx, out var fmt))
		{
			return fmt.FormatValue(val, 1900);
		}
		else
		{
			throw new FormatException();
		}
	}

	public override bool IsDBNull(int ordinal)
	{
		if (ordinal < this.columnSchema.Count && this.columnSchema[ordinal].AllowDBNull == false)
		{
			return false;
		}

		var type = this.GetExcelDataType(ordinal);
		switch (type)
		{
			case ExcelDataType.String:
				return string.IsNullOrEmpty(this.GetString(ordinal));
			case ExcelDataType.Null:
				return true;
			case ExcelDataType.Error:
				if (errorAsNull)
				{
					return true;
				}
				return false;
		}
		return false;
	}

	public override ExcelErrorCode GetFormulaError(int ordinal)
	{
		var fi = values[ordinal];
		if (fi.type == ExcelDataType.Error)
		{
			return GetErrorCode(fi.strValue);
		}
		throw new InvalidOperationException();
	}

	public override ExcelFormat? GetFormat(int ordinal)
	{
		var fi = values[ordinal];
		var idx = fi.xfIdx;

		idx = idx <= 0 ? 0 : xfMap[idx];
		if (this.formats.TryGetValue(idx, out var fmt))
		{
			return fmt;
		}
		return null;
	}

	public override int RowFieldCount => this.rowFieldCount;

	public override int MaxFieldCount => 16384;

	public override int WorksheetCount => this.sheetNames.Length;

	public override string? WorksheetName => sheetIdx < sheetNames.Length ? this.sheetNames[this.sheetIdx].Name : null;

	internal override int DateEpochYear => 1900;

	public override int RowNumber => rowIndex + 1;

	string[] stringData;

	void LoadSharedStrings(ZipArchiveEntry? entry)
	{
		if (entry == null)
		{
			this.stringData = Array.Empty<string>();
			return;
		}
		using Stream ssStream = entry.Open();
		using var reader = XmlReader.Create(ssStream);

		string ns = "";
		while (reader.Read())
		{
			if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "sst")
			{
				ns = reader.NamespaceURI;
				break;
			}
		}

		string countStr = reader.GetAttribute("uniqueCount")!;
		if (countStr == null)
		{
			stringData = Array.Empty<string>();
			return;
		}

		var count = int.Parse(countStr);

		this.stringData = new string[count];

		for (int i = 0; i < count; i++)
		{
			while (reader.Read())
			{
				if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "si")
					break;
			}

			var str = ReadString(reader);

			this.stringData[i] = str;
		}
	}

	StringBuilder stringBuilder = new StringBuilder();

	string ReadString(XmlReader reader)
	{
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
					if (reader.IsEmptyElement)
					{
						s = string.Empty;
					}
					else
					{
						reader.Read();
						if (reader.NodeType == XmlNodeType.Text || reader.NodeType == XmlNodeType.SignificantWhitespace)
						{

							s = reader.Value;
						}
						else if(reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "t")
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
		if ((uint)i >= stringData.Length)
			throw new ArgumentOutOfRangeException(nameof(i));

		return stringData[i];
	}
}
