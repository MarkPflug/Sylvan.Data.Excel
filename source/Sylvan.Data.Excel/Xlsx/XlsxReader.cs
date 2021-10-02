using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.Common;
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
		int colCount;
		int rowCount;

		Stream stream;
		XmlReader? reader;

		string currentSheetName = string.Empty;

		string[] headers;
		FieldInfo[] values;

		State state;
		bool hasRows;
		bool hasHeaders;
		IExcelSchemaProvider schema;
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

		public XlsxWorkbookReader(Stream iStream, ExcelDataReaderOptions opts)
		{
			this.colCount = 0;
			this.rowCount = 0;
			this.values = Array.Empty<FieldInfo>();
			this.headers = Array.Empty<string>();
			this.schema = opts.Schema;

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

			this.ns = sheetNS;
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
		string ns;


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
			while (!reader.IsStartElement("worksheet") && reader.Read()) ;
			this.ns = reader.NamespaceURI;
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
			int count = -1;

			if (reader == null)
			{
				this.state = State.Closed;
				throw new InvalidOperationException();
			}
			this.hasHeaders = schema.HasHeaders(currentSheetName);
			if (hasHeaders)
			{
				if (!NextRow())
				{
					return false;
				}

				count = ParseRowValues();

				if (this.headers.Length < count)
				{
					Array.Resize(ref this.headers, count);
				}

				for (int i = 0; i < count; i++)
				{
					var type = GetExcelDataType(i);
					switch (type)
					{
						case ExcelDataType.String:
							headers[i] = values[i].strValue;
							break;
						case ExcelDataType.Boolean:
							headers[i] = GetBoolean(i) ? "TRUE" : "FALSE";
							break;
						case ExcelDataType.Numeric:
							headers[i] = GetDouble(i).ToString();
							break;
						default:
							headers[i] = string.Empty;
							break;
					}
					this.GetString(i);
				}
			}
			if (!NextRow())
			{
				return false;
			}

			var c = ParseRowValues();
			count = count == -1 ? c : count;

			this.rowNumber = hasHeaders ? 0 : -1;
			this.colCount = count;
			LoadSchema();
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
				if (NextRow())
				{
					ParseRowValues();
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
						if (!int.TryParse(valueBuffer.AsSpan(0, len), NumberStyles.Integer, ci, out xfIdx))
						{
							throw new FormatException();
						}
					}
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
					reader.Read();
					switch (type)
					{
						case CellType.Numeric:
							fi.type = ExcelDataType.Numeric;
							len = reader.ReadValueChunk(valueBuffer, 0, valueBuffer.Length);
							if (len < valueBuffer.Length)
							{
								fi.numValue = double.Parse(valueBuffer.AsSpan(0, len), NumberStyles.Float, ci);
							}
							else
							{
								fi.numValue = double.Parse(reader.ReadContentAsString(), NumberStyles.Float, ci);
							}
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
							var strIdx = int.Parse(valueBuffer.AsSpan(0, len), NumberStyles.Integer, ci);
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
			return pos.Column + 1;
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

		public override ReadOnlyCollection<DbColumn> GetColumnSchema()
		{
			return columnSchema!;
		}

		ReadOnlyCollection<DbColumn>? columnSchema;

		void LoadSchema()
		{
			var cols = new List<DbColumn>();
			for (int i = 0; i < colCount; i++)
			{
				string? header = hasHeaders ? headers[i] : null;
				var col = schema.GetColumn(currentSheetName, header, i);
				var ecs = new ExcelColumn(header, i, col);
				cols.Add(ecs);
			}

			this.columnSchema = new ReadOnlyCollection<DbColumn>(cols);
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
					return fi.strValue[0] == '0' ? "FALSE" : "TRUE";
				case ExcelDataType.Numeric:
					return FormatVal(fi.xfIdx, fi.numValue);
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

		public override object GetValue(int ordinal)
		{
			var type = GetExcelDataType(ordinal);
			switch (type)
			{
				case ExcelDataType.Null:
					return DBNull.Value;
				case ExcelDataType.Boolean:
					return GetBoolean(ordinal);
				case ExcelDataType.Error:
					throw new ExcelFormulaException(RowNumber, ordinal, GetFormulaError(ordinal));
				case ExcelDataType.Numeric:
					return GetDouble(ordinal);
				case ExcelDataType.String:
					return GetString(ordinal);
			}
			throw new NotSupportedException();
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

		public override int FieldCount
		{
			get { return this.colCount; }
		}

		public override int WorksheetCount => this.sheetNames.Count;

		public override string WorksheetName => this.sheetNames[this.sheetIdx];

		internal override int DateEpochYear => 1900;

		public override int RowNumber => rowNumber;
	}

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
