using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.Common;
using System.Globalization;
using System.IO;
using System.IO.Compression;
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
					var c = int.Parse(xfsElem.GetAttribute("count"));
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
			CheckCharacters = false
		};

		const string sheetNS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
		string ns;

		public override bool NextResult()
		{
			sheetIdx++;
			var sheetName = string.Format("xl/worksheets/sheet{0}.xml", sheetIdx);

			var sheetPart = package.GetEntry(sheetName);
			if (sheetPart == null)
				return false;
			this.stream = sheetPart.Open();

			this.reader = XmlReader.Create(stream, settings);

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
			if (reader!.ReadToFollowing("row", ns))
			{
				return true;
			}

			return false;
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
					row = row * 10 + v;
				}
				return new CellPosition() { Column = col, Row = row - 1};
			}

			public static int ParseCol(ReadOnlySpan<char> str, int i)
			{
				switch (i)
				{
					case 1:
						return (int)(str[0] - 'A');
					case 2:
						return (int)
							(
								(str[0] - 'A' + 1) * 26 +
								str[1] - 'A'
							);
					case 3:
						return (int)
							(
								(str[0] - 'A' + 1) * 26 +
								(str[1] - 'A' + 1) * 26 +
								str[2] - 'A'
							);

					default:
						throw new IOException();
				}
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
			var ci = CultureInfo.InvariantCulture;
			XmlReader reader = this.reader!;
			FieldInfo[] values = this.values;

			Array.Clear(this.values, 0, this.values.Length);
			CellPosition pos;
			if (!reader.ReadToDescendant("c", ns))
			{
				return 0;
			}

			do
			{
				reader.MoveToAttribute("r");
				int len = reader.ReadValueChunk(valueBuffer, 0, valueBuffer.Length);
				pos = CellPosition.Parse(valueBuffer.AsSpan(0, len));
				if (pos.Column >= values.Length)
				{
					Array.Resize(ref values, pos.Column + 8);
					this.values = values;
				}

				CellType type = CellType.Numeric;

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
					}
					throw new NotSupportedException();
				}

				ref FieldInfo fi = ref values[pos.Column];

				bool exists = reader.MoveToAttribute("t");
				if (exists)
				{
					len = reader.ReadValueChunk(valueBuffer, 0, valueBuffer.Length);

					type = GetCellType(valueBuffer, len);
				}

				exists = reader.MoveToAttribute("s");
				int xfIdx = 0;
				if (exists)
				{
					xfIdx = reader.ReadContentAsInt();
				}
				fi.xfIdx = xfIdx;

				int strLen;
				reader.MoveToElement();

				if (reader.ReadToDescendant("v"))
				{
					reader.Read();
					fi.xfIdx = xfIdx;
					switch (type)
					{
						case CellType.Numeric:
							strLen = reader.ReadValueChunk(valueBuffer, 0, valueBuffer.Length);
							if (strLen < valueBuffer.Length)
							{
								fi.numValue = double.Parse(valueBuffer.AsSpan(0, strLen), provider: ci);
							}
							else
							{
								fi.numValue = double.Parse(reader.ReadContentAsString(), ci);
							}
							fi.type = ExcelDataType.Numeric;
							break;
						case CellType.SharedString:
							strLen = reader.ReadValueChunk(valueBuffer, 0, valueBuffer.Length);
							var strIdx = int.Parse(valueBuffer.AsSpan(0, strLen), provider: ci);
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
				do
				{
					var t = reader.NodeType;
					if ((t == XmlNodeType.EndElement || t == XmlNodeType.Element) && reader.LocalName == "c")
					{
						break;
					}
				} while (reader.Read());

			} while (reader.ReadToNextSibling("c", ns));
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
			if(rowNumber < parsedRow)
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
			var fmtIdx = this.xfMap[xfIdx];
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
			
			idx = idx == -1 ? 0 : xfMap[idx];
			if(this.formats.TryGetValue(idx, out var fmt)) {
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
			bool s;
			while (!reader.IsStartElement("sst") && reader.Read()) ;
			var ns = reader.NamespaceURI;

			string countStr = reader.GetAttribute("uniqueCount")!;
			var count = 0;
			if (countStr == null)
			{
				stringData = Array.Empty<string>();
				return;
			}
			s = reader.Read();

			count = int.Parse(countStr);
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
