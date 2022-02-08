using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace Sylvan.Data.Excel
{
	sealed class XlsxWorkbookReader : ExcelDataReader
	{
		const string SheetNS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
		const string DocRelsNS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
		const string RelationsNS = "http://schemas.openxmlformats.org/package/2006/relationships";

		readonly ZipArchive package;
		SharedStrings ss;
		Dictionary<int, ExcelFormat> formats;
		int[] xfMap;
		int sheetIdx = -1;
		int rowCount;

		bool readHiddenSheets;

		Stream stream;
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

		public XlsxWorkbookReader(Stream iStream, ExcelDataReaderOptions opts) : base(opts.Schema)
		{
			this.rowCount = -1;
			this.values = Array.Empty<FieldInfo>();

			this.refName = this.styleName = this.typeName = string.Empty;
			this.sheetNS = SheetNS;
			this.errorAsNull = opts.GetErrorAsNull;
			this.readHiddenSheets = opts.ReadHiddenWorksheets;

			this.stream = iStream;
			package = new ZipArchive(iStream, ZipArchiveMode.Read);

			var ssPart = package.GetEntry("xl/sharedStrings.xml");
			var stylePart = package.GetEntry("xl/styles.xml");

			var sheetsPart = package.GetEntry("xl/workbook.xml");
			var sheetsRelsPart = package.GetEntry("xl/_rels/workbook.xml.rels");
			if (sheetsPart == null)
				throw new InvalidDataException();

			if (ssPart == null)
			{
				ss = SharedStrings.Empty;
			}
			else
			{
				using (Stream ssStream = ssPart.Open())
				{
					ss = new SharedStrings(XmlReader.Create(ssStream));
				}
			}

			var sheetNameList = new List<SheetInfo>();
			var sheetHiddenList = new List<bool>();
			Dictionary<string, string> sheetRelMap = new Dictionary<string, string>();
			using (Stream sheetRelStream = sheetsRelsPart.Open())
			{
				var doc = new XmlDocument();
				doc.Load(sheetRelStream);
				var nsm = new XmlNamespaceManager(doc.NameTable);
				nsm.AddNamespace("r", RelationsNS);
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
				nsm.AddNamespace("x", SheetNS);
				var nodes = doc.SelectNodes("/x:workbook/x:sheets/x:sheet", nsm);
				foreach (XmlElement sheetElem in nodes)
				{
					var id = int.Parse(sheetElem.GetAttribute("sheetId"));
					var name = sheetElem.GetAttribute("name");
					var state = sheetElem.GetAttribute("state");
					var refId = sheetElem.GetAttribute("id", DocRelsNS);

					sheetHiddenList.Add(StringComparer.OrdinalIgnoreCase.Equals(state, "hidden"));
					var si = new SheetInfo(name, sheetRelMap[refId]);
					sheetNameList.Add(si);
				}
			}
			this.sheetNames = sheetNameList.ToArray();
			this.sheetHiddenFlags = sheetHiddenList.ToArray();
			if (stylePart == null)
			{
				throw new InvalidDataException();
			}
			else
			{
				using (Stream styleStream = stylePart.Open())
				{
					var doc = new XmlDocument();
					doc.Load(styleStream);
					var nsm = new XmlNamespaceManager(doc.NameTable);
					nsm.AddNamespace("x", SheetNS);
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

		public override bool IsClosed
		{
			get { return this.stream == Stream.Null; }
		}

		public override void Close()
		{
			this.stream?.Close();
			this.stream = Stream.Null;
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
			//var sheetName = $"xl/worksheets/sheet{sheetIdx + 1}.xml";

			// the relationship is recorded as an absolute path
			// but the zip entry has a relative name.
			sheetName = sheetName.TrimStart('/');
			var sheetPart = package.GetEntry(sheetName);
			if (sheetPart == null)
				return false;
			this.stream = sheetPart.Open();

			var tr = new StreamReader(this.stream, Encoding.Default, true, 0x10000);

			this.reader = XmlReader.Create(tr, settings);
			refName = this.reader.NameTable.Add("r");
			typeName = this.reader.NameTable.Add("t");
			styleName = this.reader.NameTable.Add("s");
			sheetNS = this.reader.NameTable.Add(SheetNS);

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
					if (reader.LocalName == "dimensions")
					{

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

			var c = ParseRowValues();
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
			return reader!.ReadToFollowing("row", sheetNS);
		}

		struct CellPosition
		{
			public int Column;
			public int Row;

			public static CellPosition Parse(ReadOnlySpan<char> str)
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
			CellPosition pos = default;
			if (!reader.ReadToDescendant("c", sheetNS))
			{
				return 0;
			}

			int valueCount = 0;

			do
			{
				CellType type = CellType.Numeric;
				int xfIdx = 0;
				while (reader.MoveToNextAttribute())
				{
					var n = reader.Name;
					if (ReferenceEquals(n, refName))
					{
						len = reader.ReadValueChunk(valueBuffer, 0, valueBuffer.Length);
						pos = CellPosition.Parse(valueBuffer.AsSpan(0, len));
						//pos = CellPosition.Parse(reader.Value);
						if (pos.Column >= values.Length)
						{
							Array.Resize(ref values, pos.Column + 8);
							this.values = values;
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

				static bool TryParse(ReadOnlySpan<char> span, out int value)
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
						case 'd':
							return CellType.Date;
						default:
							// TODO:
							throw new NotSupportedException();
					}
				}

				ref FieldInfo fi = ref values[pos.Column];
				fi.xfIdx = xfIdx;

				reader.MoveToElement();
				var depth = reader.Depth;

				if (reader.ReadToDescendant("v", sheetNS))
				{
					valueCount++;
					this.rowFieldCount = pos.Column + 1;
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
							fi.strValue = ss.GetString(strIdx);
							fi.type = ExcelDataType.String;
							break;
						case CellType.String:
							fi.strValue = reader.ReadContentAsString();
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

			} while (reader.ReadToNextSibling("c", sheetNS));
			this.parsedRowIndex = pos.Row;
			return valueCount == 0 ? 0 : pos.Column + 1;
		}

		enum CellType
		{
			Numeric,
			String,
			SharedString,
			Boolean,
			Error,
			Date,
		}

		public override string GetName(int ordinal)
		{
			return this.columnSchema?[ordinal].ColumnName ?? "";
		}

		public override Type GetFieldType(int ordinal)
		{
			return this.columnSchema?[ordinal].DataType ?? typeof(string);
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
			if (rowIndex < parsedRowIndex)
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
			if (this.columnSchema[ordinal].AllowDBNull == false)
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

		sealed class SharedStrings
		{
			internal static SharedStrings Empty;

			static SharedStrings()
			{
				Empty = new SharedStrings();
			}

			private SharedStrings()
			{
				this.count = 0;
				this.stringData = Array.Empty<string>();
			}

			int count;
			string[] stringData;

			public SharedStrings(XmlReader reader)
			{
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
				reader.Read();

				var count = int.Parse(countStr);
				this.count = count;
				this.stringData = new string[this.count];

				for (int i = 0; i < count; i++)
				{
					reader.ReadStartElement("si", ns);

					var empty = reader.IsEmptyElement;

					reader.ReadStartElement("t", ns);
					var str = empty ? "" : reader.ReadContentAsString();
					this.stringData[i] = str;
					if (!empty)
						reader.ReadEndElement();
					reader.ReadEndElement();
				}
			}

			public int Count
			{
				get { return this.count; }
			}

			public string GetString(int i)
			{
				if ((uint)i >= count)
					throw new ArgumentOutOfRangeException(nameof(i));

				return stringData[i];
			}
		}
	}
}
