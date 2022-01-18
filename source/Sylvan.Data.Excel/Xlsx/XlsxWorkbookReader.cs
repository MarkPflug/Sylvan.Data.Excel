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
		readonly ZipArchive package;
		SharedStrings ss;
		int sheetIdx = 0;
		int rowCount;

		Stream stream;
		XmlReader? reader;

		string currentSheetName = string.Empty;

		FieldInfo[] values;
		int rowFieldCount;
		State state;
		bool hasRows;
		bool skipEmptyRows = true; // TODO: make this an option?
		int rowNumber;
		Dictionary<int, string> sheetNames;

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

		public XlsxWorkbookReader(Stream iStream, ExcelDataReaderOptions opts) : base(opts.Schema)
		{
			this.rowCount = 0;
			this.values = Array.Empty<FieldInfo>();

			this.refName = this.styleName = this.typeName = string.Empty;

			this.stream = iStream;
			package = new ZipArchive(iStream, ZipArchiveMode.Read);

			var ssPart = package.GetEntry("xl/sharedStrings.xml");
			var stylePart = package.GetEntry("xl/styles.xml");

			var sheetsPart = package.GetEntry("xl/workbook.xml");
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

			this.sheetNames = new Dictionary<int, string>();

			using (Stream sheetsStream = sheetsPart.Open())
			{
				// quick and dirty, good enough, this doc should be small.
				var doc = new XmlDocument();
				doc.Load(sheetsStream);
				var nsm = new XmlNamespaceManager(doc.NameTable);
				nsm.AddNamespace("x", sheetNS);
				var nodes = doc.SelectNodes("/x:workbook/x:sheets/x:sheet", nsm);
				foreach (XmlElement sheetElem in nodes)
				{
					var id = int.Parse(sheetElem.GetAttribute("sheetId"));
					var name = sheetElem.GetAttribute("name");
					sheetNames.Add(id, name);
				}
			}

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
					nsm.AddNamespace("x", sheetNS);
					var nodes = doc.SelectNodes("/x:styleSheet/x:numFmts/x:numFmt", nsm);
					this.formats = ExcelFormat.CreateFormatCollection();
					foreach (XmlElement fmt in nodes)
					{
						var id = int.Parse(fmt.GetAttribute("numFmtId"));
						var str = fmt.GetAttribute("formatCode");
						var ef = new ExcelFormat(str);
						formats.Add(id, ef);
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
		Dictionary<int, ExcelFormat> formats;
		int[] xfMap;

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

		const string sheetNS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

		string refName;
		string typeName;
		string styleName;

		public override bool NextResult()
		{
			sheetIdx++;
			var sheetName = string.Format("xl/worksheets/sheet{0}.xml", sheetIdx);

			var sheetPart = package.GetEntry(sheetName);
			if (sheetPart == null)
				return false;
			this.stream = sheetPart.Open();

			var tr = new StreamReader(this.stream, Encoding.Default, true, 0x10000);

			this.reader = XmlReader.Create(tr, settings);
			refName = this.reader.NameTable.Add("r");
			typeName = this.reader.NameTable.Add("t");
			styleName = this.reader.NameTable.Add("s");

			// worksheet
			while (!reader.IsStartElement("worksheet") && reader.Read())
			{
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

			var hasHeaders = schema.HasHeaders(currentSheetName);

			LoadSchema(!hasHeaders);

			if (hasHeaders)
			{
				this.state = State.Open;
				Read();
			}

			this.rowNumber = hasHeaders ? 0 : -1;
			this.state = State.Initialized;
			return true;
		}

		bool NextRow()
		{
			return reader!.ReadToFollowing("row");
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
			rowNumber++;
			if (state == State.Open)
			{
				if (rowNumber <= parsedRow)
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
			rowNumber = -1;
			this.state = State.End;
			return false;
		}

		char[] valueBuffer = new char[64];

		int parsedRow = -1;

		int ParseRowValues()
		{
			var ci = NumberFormatInfo.InvariantInfo;
			XmlReader reader = this.reader!;
			FieldInfo[] values = this.values;
			int len;

			Array.Clear(this.values, 0, this.values.Length);
			CellPosition pos = default;
			if (!reader.ReadToDescendant("c"))
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

				if (reader.ReadToDescendant("v"))
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

			} while (reader.ReadToNextSibling("c"));
			this.parsedRow = pos.Row;
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
			if (rowNumber < parsedRow)
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
			if (rowNumber == parsedRow)
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
			return new ExcelFormulaException(ordinal, rowNumber, GetFormulaError(ordinal));
		}

		public override string GetString(int ordinal)
		{
			if (rowNumber < parsedRow)
			{
				return string.Empty;
			}
			ref var fi = ref values[ordinal];
			switch (fi.type)
			{
				case ExcelDataType.Error:
					throw Error(ordinal);
				case ExcelDataType.Boolean:
					return fi.strValue[0] == '0' ? bool.FalseString : bool.TrueString;
				case ExcelDataType.Numeric:
					return FormatVal(fi.xfIdx, fi.numValue);
				case ExcelDataType.DateTime:
					return IsoDate.ToStringIso(fi.dtValue);
			}
			return fi.strValue;
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
			return GetExcelDataType(ordinal) == ExcelDataType.Null;
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

		public override int WorksheetCount => this.sheetNames.Count;

		public override string? WorksheetName => this.sheetNames.ContainsKey(this.sheetIdx) ? this.sheetNames[this.sheetIdx] : null;

		internal override int DateEpochYear => 1900;

		public override int RowNumber => rowNumber;

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

			const string ssNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

			public SharedStrings(XmlReader reader)
			{
				while (!reader.IsStartElement("sst") && reader.Read()) ;

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
					reader.ReadStartElement("si");

					var empty = reader.IsEmptyElement;

					reader.ReadStartElement("t");
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
