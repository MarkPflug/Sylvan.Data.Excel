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
		IExcelSchemaProvider schema;
		int rowNumber;

		struct FieldInfo
		{
			public ExcelDataType type;
			public string strValue;
			public double numValue;
		}

		public override int RowCount => rowCount;

		public XlsxWorkbookReader(Stream iStream, ExcelDataReaderOptions opts)
		{
			this.colCount = 0;
			this.rowCount = 0;
			this.headers = new string[8];
			this.values = Array.Empty<FieldInfo>();
			this.schema = opts.Schema;

			this.stream = iStream;
			package = new ZipArchive(iStream, ZipArchiveMode.Read);
			
			var ssPart = package.GetEntry("xl/sharedStrings.xml");
			if (ssPart == null)
				throw new InvalidDataException();

			using (Stream ssStream = ssPart.Open())
			{
				ss = new SharedStrings(XmlReader.Create(ssStream));
			}
			this.ns = sheetNS;
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
			rowNumber = 0;
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

			if (schema.HasHeaders(currentSheetName))
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
			//public int Column;
			//public int Row;

			public static int Parse(ReadOnlySpan<char> str)
			{
				int c = -1;
				for (int i = 0; i < str.Length; i++)
				{
					var cc = str[i];
					var v = cc - 'A';
					if ((uint)v < 26u)
					{
						c = ((c + 1) * 26) + v;
					}
					else
					{
						break;
					}
				}
				return c;
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

		int ParseRowValues()
		{
			var ci = CultureInfo.InvariantCulture;
			XmlReader reader = this.reader!;
			FieldInfo[] values = this.values;

			Array.Clear(this.values, 0, this.values.Length);
			int col;
			if (!reader.ReadToDescendant("c", ns))
			{
				return 0;
			}

			do
			{
				reader.MoveToAttribute("r");
				int len = reader.ReadValueChunk(valueBuffer, 0, valueBuffer.Length);
				col = CellPosition.Parse(valueBuffer.AsSpan(0, len));
				if (col >= values.Length)
				{
					Array.Resize(ref values, col + 8);
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

				bool exists = reader.MoveToAttribute("t");
				if (exists)
				{
					len = reader.ReadValueChunk(valueBuffer, 0, valueBuffer.Length);

					type = GetCellType(valueBuffer, len);
				}

				int strLen;
				reader.MoveToElement();

				if (reader.ReadToDescendant("v"))
				{
					reader.Read();
					ref FieldInfo fi = ref values[col];

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
			return col + 1;
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
				string? header = headers[i];
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
			throw new InvalidCastException();
		}

		ExcelFormulaException Error(int ordinal)
		{
			return new ExcelFormulaException(ordinal, rowNumber, GetFormulaError(ordinal));
		}

		public override string GetString(int ordinal)
		{
			ref var fi = ref values[ordinal];
			switch (fi.type)
			{
				case ExcelDataType.Error:
					throw Error(ordinal);
				case ExcelDataType.Boolean:
					return fi.strValue[0] == '0' ? "FALSE" : "TRUE";
				case ExcelDataType.Numeric:
					return fi.numValue.ToString();
			}
			return fi.strValue;
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
			throw new NotImplementedException();
		}

		public override int FieldCount
		{
			get { return this.colCount; }
		}

		public override int WorksheetCount => throw new NotImplementedException();
		public override string WorksheetName => throw new NotImplementedException();

		internal override int DateEpochYear => 1900;

		public override int RowNumber => rowNumber;
	}

	sealed class SharedStrings
	{
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
			if ((uint) i >= count)
				throw new ArgumentOutOfRangeException(nameof(i));

			return stringData[i];
		}
	}
}
