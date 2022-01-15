using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace Sylvan.Data.Excel;

sealed class XlsbWorkbookReader : ExcelDataReader
{
	readonly ZipArchive package;
	SharedStrings ss;
	int sheetIdx = 0;
	int rowCount;

	Stream stream;

	string currentSheetName = string.Empty;

	FieldInfo[] values;
	int rowFieldCount;
	State state;
	bool hasRows = false;
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

	public XlsbWorkbookReader(Stream iStream, ExcelDataReaderOptions opts) : base(opts.Schema)
	{
		this.rowCount = 0;
		this.values = Array.Empty<FieldInfo>();

		this.refName = this.styleName = this.typeName = string.Empty;

		this.stream = iStream;
		package = new ZipArchive(iStream, ZipArchiveMode.Read);

		var ssPart = package.GetEntry("xl/sharedStrings.bin");
		var stylePart = package.GetEntry("xl/styles.bin");

		var sheetsPart = package.GetEntry("xl/workbook.bin");
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
				ss = new SharedStrings(ssStream);
			}
		}

		this.sheetNames = new Dictionary<int, string>();

		using (Stream sheetsStream = sheetsPart.Open())
		{
			var rr = new RecordReader(sheetsStream);
			var atEnd = false;
			while (!atEnd)
			{
				rr.NextRecord();
				switch (rr.RecordType)
				{

					case RecordType.BundleBegin:

						while (true)
						{
							rr.NextRecord();
							if (rr.RecordType == RecordType.BundleSheet)
							{
								var vis = rr.GetInt32(0);
								var id = rr.GetInt32(4);
								var rel = rr.GetString(8, out int next);
								var name = rr.GetString(next);
								this.sheetNames.Add(id, name);
							}
							else
							if (rr.RecordType == RecordType.BundleEnd)
							{
								break;
							}

						}
						break;

					case RecordType.BookEnd:
						atEnd = true;
						break;
				}
				//rr.DebugInfo("sheets");
			}
		}

		if (stylePart == null)
		{
			throw new InvalidDataException();
		}
		else
		{
			ReadStyle(stylePart);
		}

		NextResult();
	}

	static readonly Dictionary<int, ExcelFormat> EmptyFormats = new Dictionary<int, ExcelFormat>();

	Dictionary<int, ExcelFormat> formats = EmptyFormats;
	int[] xfMap = Array.Empty<int>();

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

	string refName;
	string typeName;
	string styleName;

	RecordReader? reader;

	public override bool NextResult()
	{
		sheetIdx++;
		if (sheetIdx > this.sheetNames.Count)
			return false;

		var sheetName = $"xl/worksheets/sheet{sheetIdx}.bin";

		var sheetPart = package.GetEntry(sheetName);
		if (sheetPart == null)
			return false;
		this.stream = sheetPart.Open();

		this.reader = new RecordReader(this.stream);
		var rr = this.reader;
		this.rowFieldCount = 0;

		return InitializeSheet();
	}

	bool InitializeSheet()
	{
		this.hasRows = true;
		this.state = State.Initializing;

		if (reader == null)
		{
			this.state = State.Closed;
			throw new InvalidOperationException();
		}

		while (true)
		{
			reader.NextRecord();

			if (reader.RecordType == RecordType.Dimension)
			{
				var rowLast = reader.GetInt32(4);
				var colLast = reader.GetInt32(12);
				if(this.values.Length < colLast + 1)
				{
					Array.Resize(ref values, colLast + 1);
				}
			}
			if (reader.RecordType == RecordType.DataStart)
			{
				break;
			}
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
			while (true)
			{
				var c = ParseRowValues();
				if (c < 0)
					return false;
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
		if (reader == null)
			throw new InvalidOperationException();

		Array.Clear(values, 0, values.Length);

		while(reader.RecordType != RecordType.Row)
		{
			reader.NextRecord();
			if (reader.RecordType == RecordType.DataEnd)
				return -1;
		}

		var rowIdx = reader.GetInt32(0);
		var ifx = reader.GetInt32(4);

		reader.NextRecord();
		int count = 0;
		while (reader.RecordType != RecordType.Row)
		{
			reader.DebugInfo("data");
			switch (reader.RecordType)
			{
				case RecordType.CellRK:
					var col = reader.GetInt32(0);
					var sf = reader.GetInt32(4) & 0xffffff;
					var rk = reader.GetInt32(8);

					var mul100 = (rk & 1) == 1;
					var mode = (rk & 2) == 2;

					var val = rk >> 2 & 0x3fffffff;

					double d = mode
						? val
						: BitConverter.Int64BitsToDouble(((long)val) << 34);
					ref var fi = ref values[col];

					fi.type = ExcelDataType.Numeric;
					fi.numValue = d;
					count++;
					break;
			}
			if (reader.RecordType == RecordType.DataEnd)
			{
				return -1;
			}				
			reader.NextRecord();
		}

		this.parsedRow = rowIdx;
		return count;
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
				return fi.strValue[0] == '0' ? "FALSE" : "TRUE";
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

	public override string WorksheetName => this.sheetNames[this.sheetIdx];

	internal override int DateEpochYear => 1900;

	public override int RowNumber => rowNumber;

	void ReadStyle(ZipArchiveEntry part)
	{
		using (Stream styleStream = part.Open())
		{
			var rr = new RecordReader(styleStream);
			bool atEnd = false;
			rr.NextRecord();
			if (rr.RecordType != RecordType.StyleBegin)
			{
				throw new InvalidDataException();
			}
			int[] ixf;
			Dictionary<int, ExcelFormat> formats;

			while (!atEnd)
			{
				rr.NextRecord();
				int count;
				switch (rr.RecordType)
				{
					case RecordType.StyleEnd:
						atEnd = true;
						break;
					case RecordType.CellXFStart:
						count = rr.GetInt32();
						ixf = new int[count];
						for (int i = 0; i < count; i++)
						{
							rr.NextRecord();
							if (rr.RecordType != RecordType.XF)
							{
								throw new InvalidDataException();
							}
							ixf[i] = rr.GetInt16(2);
						}
						rr.NextRecord();
						if (rr.RecordType != RecordType.CellXFEnd)
						{
							throw new InvalidDataException();
						}
						this.xfMap = ixf;
						break;
					case RecordType.FmtStart:
						count = rr.GetInt32();
						formats = new Dictionary<int, ExcelFormat>(count);
						for (int i = 0; i < count; i++)
						{
							rr.NextRecord();
							if (rr.RecordType != RecordType.Fmt)
							{
								throw new InvalidDataException();
							}
							var id = rr.GetInt16(0);
							var fmtStr = rr.GetString(2);
							formats.Add(id, new ExcelFormat(fmtStr));
						}
						rr.NextRecord();
						if (rr.RecordType != RecordType.FmtEnd)
						{
							throw new InvalidDataException();
						}
						this.formats = formats;
						break;
				}
			}
		}
	}

	sealed class RecordReader
	{
		readonly Stream stream;
		BinaryReader reader;
		byte[] data;

		[Conditional("DEBUG")]
		internal void DebugInfo(string header)
		{
			Debug.WriteLine(header + ": " + this.RecordType + " " + this.recordLen + " " + Encoding.ASCII.GetString(data, 0, this.recordLen).Replace('\0', '_'));
		}

		public RecordReader(Stream stream)
		{
			this.stream = stream;
			this.reader = new BinaryReader(stream);
			this.data = new byte[0x1000];
		}

		internal RecordType RecordType => type;
		internal int RecordLen => recordLen;

		RecordType type;
		int recordLen;

		public int GetInt32(int offset = 0)
		{
			return BitConverter.ToInt32(data, offset);
		}

		public short GetInt16(int offset = 0)
		{
			return BitConverter.ToInt16(data, offset);
		}

		public string GetString(int offset)
		{
			var len = BitConverter.ToInt32(data, offset);
			return Encoding.Unicode.GetString(data, offset + 4, len * 2);
		}

		public string GetString(int offset, out int end)
		{
			var len = BitConverter.ToInt32(data, offset);
			end = offset + 4 + len * 2;
			return Encoding.Unicode.GetString(data, offset + 4, len * 2);
		}

		public bool NextRecord()
		{
			type = reader.ReadRecordType();
			recordLen = reader.ReadRecordLen();
			if (recordLen > data.Length)
			{
				// TODO: allocate with some overhead?
				// maybe round to next power of two.
				Array.Resize(ref data, recordLen);
			}
			var count = reader.Read(data, 0, recordLen);
			if (count != recordLen)
				throw new Exception();
			return true;
		}
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

		public SharedStrings(Stream stream)
		{
			byte[] buffer = new byte[0x1000];

			var br = new BinaryReader(stream);
			var type = br.ReadRecordType();
			if (type != RecordType.SSTBegin)
				throw new InvalidDataException();

			var len = br.ReadRecordLen();
			if (len != 8) throw new InvalidDataException();

			int totalCount = br.ReadInt32();
			int count = br.ReadInt32();

			stringData = new string[count];

			for (int i = 0; i < count; i++)
			{
				type = br.ReadRecordType();
				if (type != RecordType.SSTItem)
					throw new InvalidDataException();

				len = br.ReadRecordLen();

				var flags = br.ReadByte();
				if (flags == 0)
				{
					len = br.ReadInt32();
					if (len > 0x7fff) throw new InvalidDataException();
					if (len > buffer.Length * 2)
					{
						Array.Resize(ref buffer, len * 2);
					}
					br.Read(buffer, 0, len * 2);
					var str = Encoding.Unicode.GetString(buffer, 0, len * 2);
					stringData[i] = str;
					// "seek" past any remaining record.
					// the stream can't seek because it is a deflate stream.
					var remains = len - (5 + len * 2);
					if (remains > 0)
					{
						stream.Read(buffer, 0, remains);
					}
				}
				else
				{
					throw new NotImplementedException();
				}
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

enum RecordType
{
	Row = 0,
	CellRK = 2,
	SSTItem = 19,
	Fmt = 44,
	XF = 47,
	BundleBegin = 143,
	BundleEnd = 144,
	BundleSheet = 156,
	BookBegin = 131,
	BookEnd = 132,
	Dimension = 148,
	SSTBegin = 159,
	StyleBegin = 278,
	StyleEnd = 279,
	CellXFStart = 617,
	CellXFEnd = 618,
	FmtStart = 615,
	FmtEnd = 616,
	SheetStart = 129,
	SheetEnd = 130,
	DataStart = 145,
	DataEnd = 146,
}

static class BinaryExtensions
{
	public static RecordType ReadRecordType(this BinaryReader br)
	{
		var b = br.ReadByte();
		if ((b & 0x80) == 0)
		{
			return (RecordType)b;
		}
		return (RecordType)(b & 0x7f | (br.ReadByte() << 7));
	}

	public static int ReadRecordLen(this BinaryReader br)
	{
		int accum = 0;
		int shift = 0;
		for (int i = 0; i < 4; i++, shift += 7)
		{
			var b = br.ReadByte();
			accum = (b & 0x7f) << shift;
			if ((b & 0x80) == 0)
				break;
		}
		return accum;
	}
}
