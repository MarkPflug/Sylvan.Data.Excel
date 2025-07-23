#nullable enable
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;

#if !SPAN
using ReadonlyCharSpan = System.String;
#else
using ReadonlyCharSpan = System.ReadOnlySpan<char>;
using CharSpan = System.Span<char>;
#endif

namespace Sylvan.Data.Excel;

sealed partial class XlsxWorkbookReader : ExcelDataReader
{
	readonly ZipArchive package;

	Stream sheetStream;
	XmlReader? reader;

	StringBuilder? stringBuilder;

	bool hasRows;

	// the number of fields in the parsedRowIndex.
	int curFieldCount = -1;
	int parsedRowIndex = -1;

	public override ExcelWorkbookType WorkbookType => ExcelWorkbookType.ExcelXml;

	const string DefaultWorkbookPartName = "xl/workbook.xml";
	readonly ZipArchiveEntry? sstPart;
	XmlReader? sstReader;
	int sstIdx = -1;

	public override void Close()
	{
		this.reader?.Close();
		this.sstReader?.Close();
		base.Close();
	}

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

		sstPart = package.FindEntry(sharedStringsPartName);
		var stylePart = package.FindEntry(stylesPartName);

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

					if (!sheetRelMap.TryGetValue(refId, out var part))
					{
						continue;
					}

					var hidden = StringComparer.OrdinalIgnoreCase.Equals(state, "hidden");
					var si = new SheetInfo(name, part, hidden);
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
						if (fmtIdStr != null)
						{
							var id = int.TryParse(fmtIdStr, out int val) ? val : 0;
							xfMap[idx] = id;
						}
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

	public override bool IsRowHidden
	{
		get
		{
			return this.isRowHidden && this.rowIndex == this.parsedRowIndex;
		}
	}

	private protected override ref readonly FieldInfo GetFieldValue(int ordinal)
	{
		if (rowIndex < this.parsedRowIndex)
		{
			return ref FieldInfo.Null;
		}
		return ref base.GetFieldValue(ordinal);
	}

	private protected override bool OpenWorksheet(int sheetIdx)
	{
		var sheetName = sheetInfos[sheetIdx].Part;
		var sheetPart = package.FindEntry(sheetName);
		if (sheetPart == null)
			return false;

		this.sheetStream = sheetPart.Open();

		var tr = new StreamReader(this.sheetStream, Encoding.UTF8, true, 0x10000);

		var settings = new XmlReaderSettings
		{
			CheckCharacters = false,
			CloseInput = true,
			ValidationType = ValidationType.None,
			ValidationFlags = System.Xml.Schema.XmlSchemaValidationFlags.None,
#if SPAN
			NameTable = new SheetNameTable(),
#endif
		};

		this.reader = XmlReader.Create(tr, settings);
		this.rowIndex = 0;
		this.rowFieldCount = this.curFieldCount = 0;
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
				else
				if (reader.LocalName == "cols")
				{
					if (reader.IsEmptyElement)
						continue;
					while (reader.Read())
					{
						if (reader.NodeType == XmlNodeType.Element)
						{
							if (reader.LocalName == "col")
							{
								int min = 0, max = 0;
								bool hidden = false;
								while (reader.MoveToNextAttribute())
								{
									var name = reader.LocalName;
									if (name == "hidden")
									{
										hidden = ReadBooleanValue(reader, buffer);
									}
									else
									if (name == "min")
									{
										if (!TryReadIntValue(reader, buffer, out min))
										{
											// TODO ? This means there was an attribute but it didn't contain a valid int value
											// which I think we can just treat as zero.
										}
									}
									else
									if (name == "max")
									{
										if (!TryReadIntValue(reader, buffer, out max))
										{
											// TODO ?
										}
									}
								}

								if (min <= max && min != 0 && max != 0)
								{
									for (int i = min; i <= max; i++)
									{
										colHidden[i-1] = hidden;
									}
								}
							}
						}
						else
						if (reader.NodeType == XmlNodeType.EndElement)
						{
							break;
						}
					}
				}
				else
				if (reader.LocalName == "sheetData")
				{
					break;
				}
			}
		}

		this.hasRows = InitializeSheet();
		this.sheetIdx = sheetIdx;
		return true;
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

		return OpenWorksheet(sheetIdx);
	}

	bool InitializeSheet()
	{
		this.state = State.Initializing;
		this.fieldCount = 0;
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
		else
		{
			this.curFieldCount = rowFieldCount;
		}

		this.state = State.Initialized;
		this.rowIndex = LoadSchema() ? -1 : 0;

		return true;
	}

	// a buffer used to read values that must be materialized when read.
	char[] buffer = new char[16];
	// a buffer used to hold values that can be materialized lazily when the field is accessed.
	char[] valuesBuffer = Array.Empty<char>();

	const int ValueBufferElementSize = 64;

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

	bool NextRow()
	{
		var reader = this.reader;
		if (reader == null) return false;

		var ci = NumberFormatInfo.InvariantInfo;
		if (ReadToFollowing(reader, "row"))
		{
			this.parsedRowIndex = -1;
			this.isRowHidden = false;
			while (reader.MoveToNextAttribute())
			{
				if (reader.LocalName == "r")
				{
					int row;
#if SPAN
					var len = reader.ReadValueChunk(buffer, 0, buffer.Length);
					if (len < buffer.Length && TryParse(buffer.AsSpan(0, len), out row))
					{
					}
					else
					{
						row = 0;
					}
#else
					var str = reader.Value;
					if (!int.TryParse(str, NumberStyles.Integer, ci, out row))
					{
						row = 0;
					}
#endif
					this.parsedRowIndex = row - 1;

				}
				else
				if (reader.LocalName == "hidden")
				{
					this.isRowHidden = ReadBooleanValue(reader, buffer);
				}
			}
			reader.MoveToElement();
			return true;
		}
		return false;
	}

	static bool ReadBooleanValue(XmlReader reader, char[] buffer)
	{
		var len = reader!.ReadValueChunk(buffer, 0, 1);
		if (len == 1)
		{
			var c = buffer[0];
			switch (c)
			{
				case '0':
				case 'f':
				case 'F':
					return false;
				case '1':
				case 't':
				case 'T':
					return true;
			}
		}
		return false;// empty value.
	}

	static bool TryReadIntValue(XmlReader reader, char[] buffer, out int value)
	{
		var len = reader!.ReadValueChunk(buffer, 0, buffer.Length);
		if (len > 0)
		{
			if (TryParse(buffer.AsSpan().ToParsable(0, len), out value))
			{
				return true;
			}
		}
		value = 0;
		return false;
	}

	struct CellPosition
	{
		public int Column;
		public int Row;

		public static bool TryParse(ReadonlyCharSpan str, out CellPosition pos)
		{
			pos = default;
			if (str.Length == 0)
				return false;

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
		colCacheIdx = 0;
	start:
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
				if (this.readHiddenRows == false && this.IsRowHidden)
				{
					rowIndex++;
					continue;
				}
				if (c == 0)
				{
					if (this.ignoreEmptyTrailingRows)
					{
						// handles trailing empty rows.
						continue;
					}

					this.rowFieldCount = 0;
				}
				if (rowIndex < parsedRowIndex)
				{
					this.curFieldCount = c;
					this.rowFieldCount = 0;
				}
				return true;
			}
		}
		else
		if (state == State.Initialized && hasRows)
		{
			this.state = State.Open;
			if (rowIndex == 1) goto start;
			if (rowIndex == parsedRowIndex && curFieldCount >= 0)
			{
				this.rowFieldCount = curFieldCount;
				this.curFieldCount = -1;
			}
			return true;
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
				var n = reader.LocalName;
				if (n == "r")
				{
					len = reader.ReadValueChunk(buffer, 0, buffer.Length);

					if (CellPosition.TryParse(buffer.AsSpan().ToParsable(0, len), out var pos))
					{
						col = pos.Column;
					}
					else
					{
						// if the cell ref is unparsable, Excel seems to treat it as missing.
					}
				}
				else
				if (n == "t")
				{
					len = reader.ReadValueChunk(buffer, 0, buffer.Length);
					type = GetCellType(buffer, len);
				}
				else
				if (n == "s")
				{
					len = reader.ReadValueChunk(buffer, 0, buffer.Length);
					if (!TryParse(buffer.AsSpan().ToParsable(0, len), out xfIdx))
					{
						throw new FormatException();
					}
				}
			}

			if (col >= values.Length)
			{
				var newLen = col + 8;

				Array.Resize(ref values, newLen);
				this.values = values;
				Array.Resize(ref valuesBuffer, newLen * ValueBufferElementSize);
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
					fi.type = FieldType.String;
					valueCount++;
					this.rowFieldCount = col + 1;
				}
				else
				{
					fi.strValue = string.Empty;
					fi.type = FieldType.Null;
				}
			}
			else
			if (ReadToDescendant(reader, "v"))
			{
				if (reader.IsEmptyElement)
				{
					fi.type = FieldType.Null;
					fi.valueLen = 0;
				}
				else
				{
					reader.Read();

					int ReadValue(int col)
					{
						return reader.ReadValueChunk(valuesBuffer, col * ValueBufferElementSize, ValueBufferElementSize);
					}

					if (reader.NodeType == XmlNodeType.Text)
					{
						switch (type)
						{
							case CellType.Numeric:
								fi.type = FieldType.Numeric;
								fi.valueLen = ReadValue(col);
								break;
							case CellType.Date:
								fi.type = FieldType.DateTime;
								fi.valueLen = ReadValue(col);
								break;
							case CellType.SharedString:
								if (reader.NodeType == XmlNodeType.Text)
								{
									fi.type = FieldType.SharedString;
									fi.valueLen = ReadValue(col);
								}
								else
								{
									// this handles an edge-case where the field is a shared string,
									// but the index is empty.
									fi.strValue = string.Empty;
									fi.type = FieldType.String;
								}
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
									fi.type = FieldType.String;
								}
								else
								{
									fi.strValue = string.Empty;
									fi.type = FieldType.Null;
								}
								break;
							case CellType.InlineString:
								fi.strValue = ReadString(reader);
								fi.type = FieldType.String;
								if (fi.strValue.Length == 0)
								{
									fi.type = FieldType.Null;
								}
								break;
							case CellType.Boolean:
								fi.type = FieldType.Boolean;
								fi.valueLen = ReadValue(col);
								//fi = new FieldInfo(valueBuffer[0] != '0');
								break;
							case CellType.Error:
								fi.type = FieldType.Error;
								fi.valueLen = ReadValue(col);
								//fi = new FieldInfo(GetErrorCode(valueBuffer.AsSpan(0, len)));
								break;
							default:
								throw new InvalidDataException();
						}
						if (fi.type != FieldType.Null)
						{
							valueCount++;
							this.rowFieldCount = col + 1;
						}
					}
				}
			}

			while (reader.Depth > depth)
			{
				reader.Read();
			}

		} while (ReadToNextSibling(reader, "c"));
		return valueCount == 0 ? 0 : col + 1;
	}

	static bool Equal(ReadonlyCharSpan l, ReadonlyCharSpan r)
	{
		if (l.Length != r.Length) return false;
		for (int i = 0; i < l.Length; i++)
		{
			if (l[i] != r[i])
				return false;
		}
		return true;
	}

	static ExcelErrorCode GetErrorCode(ReadonlyCharSpan str)
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
			start:
				if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "rPh")
				{
					SkipSubtree(reader);
					// after skipping the subtree, the reader will already be positioned on the next element
					// so we need to avoid calling Read again.
					goto start;
				}
				else
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

	private protected override string GetSharedString(int idx)
	{
		if (this.sstIdx < idx)
		{
			if (!LoadSharedString(idx))
			{
				throw new InvalidDataException();
			}
		}
		return sst[idx];
	}

	bool LoadSharedString(int i)
	{
		var reader = this.sstReader;
		if (reader == null)
		{
			var sstStream = sstPart!.Open();
			var settings = new XmlReaderSettings
			{
				CloseInput = true,
				CheckCharacters = false,
#if SPAN
				// name table optimization requires ROS
				NameTable = new SharedStringsNameTable(),
#endif
			};

			reader = this.sstReader = XmlReader.Create(sstStream, settings);
			// advance to the content
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
				var estimatedCount = (int)(sstPart.Length / 24);
				count = Math.Max(1, estimatedCount);
			}
			if (count > 128)
				count = 128;
			this.sst = new string[count];
		}

		while (i > sstIdx)
		{
			if (reader.Read())
			{
				if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "si")
				{
					var str = ReadString(reader);
					sstIdx++;
					if (sstIdx >= sst.Length)
					{
						Array.Resize(ref sst, sst.Length * 2);
					}
					sst[sstIdx] = str;

				}
			}
			else
			{
				// a cell with an SST value reference out of bounds.
				// this exception type is probably wrong
				return false;
			}
		}

		return true;
	}
}
