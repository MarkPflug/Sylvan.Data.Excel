﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace Sylvan.Data.Excel;

sealed class XlsbWorkbookReader : ExcelDataReader
{
	readonly ZipArchive package;
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
		public ExcelErrorCode err;
		public bool b;
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


		var stylePart = package.GetEntry("xl/styles.bin");

		var sheetsPart = package.GetEntry("xl/workbook.bin");
		if (sheetsPart == null)
			throw new InvalidDataException();

		stringData = ReadSharedStrings();

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
			//reader.DebugInfo("header");
			if (reader.RecordType == RecordType.Dimension)
			{
				var rowLast = reader.GetInt32(4);
				var colLast = reader.GetInt32(12);
				if (this.values.Length < colLast + 1)
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

	string[] stringData;

	string[] ReadSharedStrings()
	{
		var ssPart = package.GetEntry("xl/sharedStrings.bin");
		if (ssPart == null)
		{
			return Array.Empty<string>();
		}
		using (var stream = ssPart.Open())
		{
			var reader = new RecordReader(stream);

			reader.NextRecord();
			if (reader.RecordType != RecordType.SSTBegin)
				throw new InvalidDataException();

			int totalCount = reader.GetInt32(0);
			int count = reader.GetInt32(4);

			var ss = new string[count];

			for (int i = 0; i < count; i++)
			{
				reader.NextRecord();
				if (reader.RecordType != RecordType.SSTItem)
				{
					reader.DebugInfo("fail");
					throw new InvalidDataException();
				}

				var flags = reader.GetByte(0);
				if (flags == 0)
				{
					var str = reader.GetString(1);
					ss[i] = str;
				}
				else
				{
					throw new NotImplementedException();
				}
			}
			return ss;
		}
	}

	public override bool Read()
	{
		rowNumber++;

		if(rowNumber == 553)
		{
			;
		}
		if (state == State.Open)
		{
			if (rowNumber <= parsedRow)
			{
				return true;
			}

			while (true)
			{
				var c = ParseRowValues();
				//Debug.Write("Read " + c);
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

	int parsedRow = -1;

	int ParseRowValues()
	{
		if (reader == null)
			throw new InvalidOperationException();

		Array.Clear(values, 0, values.Length);

		while (reader.RecordType != RecordType.Row)
		{
			if (reader.RecordType == RecordType.DataEnd)
				return -1;
			reader.NextRecord();
		}

		var rowIdx = reader.GetInt32(0);
		var ifx = reader.GetInt32(4);

		//reader.DebugInfo("parse " + rowIdx);

		reader.NextRecord();
		int count = 0;
		int notNull = 0;

		ExcelDataType type = 0;

		while (reader.RecordType != RecordType.Row)
		{
			if (reader.RecordType == RecordType.DataEnd)
			{
				if (count == 0)
				{
					return -1;
				}
				else
				{
					break;
				}
			}

			//reader.DebugInfo("data");
			switch (reader.RecordType)
			{
				case RecordType.CellRK:
					{
						var col = reader.GetInt32(0);
						var sf = reader.GetInt32(4) & 0xffffff;
						var rk = reader.GetInt32(8);

						var mul100 = (rk & 1) == 1;
						var mode = (rk & 2) == 2;

						var val = rk >> 2 & 0x3fffffff;

						double d = mode
							? val
							: BitConverter.Int64BitsToDouble(((long)val) << 34);
						d = mul100 ? d / 100d : d;
						ref var fi = ref values[col];

						fi.type = ExcelDataType.Numeric;
						fi.numValue = d;
						fi.xfIdx = sf;
						notNull++;
						count = col + 1;
					}
					break;
				case RecordType.CellNum:
					{
						var col = reader.GetInt32(0);
						var sf = reader.GetInt32(4) & 0xffffff;
						double d = reader.GetDouble(8);

						ref var fi = ref values[col];
						fi.type = ExcelDataType.Numeric;
						fi.numValue = d;
						fi.xfIdx = sf;
						count = col + 1;
						notNull++;
					}
					break;
				case RecordType.CellBlank:
				case RecordType.CellBool:
				case RecordType.CellError:
				case RecordType.CellFmlaBool:
				case RecordType.CellFmlaNum:
				case RecordType.CellFmlaError:
				case RecordType.CellFmlaString:
				case RecordType.CellIsst:
				case RecordType.CellSt:
					{
						var col = reader.GetInt32(0);
						var sf = reader.GetInt32(4) & 0xffffff;
						ref var fi = ref values[col];

						switch (reader.RecordType)
						{
							case RecordType.CellBlank:
								type = ExcelDataType.Null;
								break;
							case RecordType.CellBool:
							case RecordType.CellFmlaBool:
								type = ExcelDataType.Boolean;
								fi.b = reader.GetByte(8) != 0;
								notNull++;
								break;
							case RecordType.CellError:
							case RecordType.CellFmlaError:
								type = ExcelDataType.Error;
								fi.err = (ExcelErrorCode)reader.GetByte(8);
								notNull++;
								break;
							case RecordType.CellIsst:
								type = ExcelDataType.String;
								var sstIdx = reader.GetInt32(8);
								fi.strValue = stringData[sstIdx];
								notNull++;
								break;
							case RecordType.CellSt:
							case RecordType.CellFmlaString:
								type = ExcelDataType.String;
								fi.strValue = reader.GetString(8);
								notNull++;
								break;
							case RecordType.CellFmlaNum:
								type = ExcelDataType.Numeric;
								fi.numValue = reader.GetDouble(8);
								notNull++;
								break;
						}


						fi.type = type;
						fi.xfIdx = sf;
						count = col + 1;
					}
					break;
				default:
					//reader.DebugInfo("unk");
					break;
			}

			reader.NextRecord();
		}

		this.parsedRow = rowIdx;
		this.rowFieldCount = count;
		return notNull;
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

	public override bool GetBoolean(int ordinal)
	{
		var fi = this.values[ordinal];
		switch (fi.type)
		{
			case ExcelDataType.Boolean:
				return fi.b;
			case ExcelDataType.Numeric:
				return this.GetDouble(ordinal) != 0;
			case ExcelDataType.String:
				return bool.TryParse(fi.strValue, out var b)
					? b
					: throw new FormatException();
			case ExcelDataType.Error:
				var code = values[ordinal].err;
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
				return fi.b ? bool.TrueString : bool.FalseString;
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
			return values[ordinal].err;
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
			this.formats = ExcelFormat.CreateFormatCollection();

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
						break;
				}
			}
		}
	}

	sealed class RecordReader
	{
		const int DefaultBufferSize = 0x10000;

		readonly Stream stream;
		byte[] data;
		int pos = 0;
		int s = 0;
		int end = 0;

		RecordType type;
		int recordLen;

		[Conditional("DEBUG")]
		internal void DebugInfo(string header)
		{
			Debug.WriteLine(header + ": " + this.RecordType + " " + this.recordLen + " " + Encoding.ASCII.GetString(data, s, this.recordLen).Replace('\0', '_'));
		}

		public RecordReader(Stream stream)
		{
			this.stream = stream;
			this.data = new byte[DefaultBufferSize];
		}

		internal RecordType RecordType => type;
		internal int RecordLen => recordLen;


		public int GetInt32(int offset = 0)
		{
			return BitConverter.ToInt32(data, s + offset);
		}

		public double GetDouble(int offset = 0)
		{
			return BitConverter.ToDouble(data, s + offset);
		}

		public short GetInt16(int offset = 0)
		{
			return BitConverter.ToInt16(data, s + offset);
		}

		public short GetByte(int offset = 0)
		{
			return data[s + offset];
		}

		public string GetString(int offset)
		{
			var len = BitConverter.ToInt32(data, s + offset);
			return Encoding.Unicode.GetString(data, s + offset + 4, len * 2);
		}

		public string GetString(int offset, out int end)
		{
			var len = BitConverter.ToInt32(data, s + offset);
			end = offset + 4 + len * 2;
			return Encoding.Unicode.GetString(data, s + offset + 4, len * 2);
		}

		void FillBuffer(int requiredLen)
		{
			Debug.Assert(pos <= end);

			if(this.data.Length < requiredLen)
			{
				Array.Resize(ref this.data, requiredLen);

			}

			if (pos != end)
			{
				// TODO: make sure overlapped copy is safe here
				Buffer.BlockCopy(data, pos, data, 0, end - pos);			
			}
			end = end - pos;
			pos = 0;

			while (end < requiredLen)
			{
				var l = stream.Read(data, end, data.Length - end);
				if (l == 0)
					throw new EndOfStreamException();

				end += l;
			}
			Debug.Assert(pos <= end);
		}

		RecordType ReadRecordType()
		{
			Debug.Assert(pos <= end);

			if (pos >= end)
			{
				FillBuffer(1);
				if (pos >= end)
					throw new EndOfStreamException();
			}

			var b = data[pos++];
			if ((b & 0x80) == 0)
			{
				return (RecordType)b;
			}

			var type = (RecordType)(b & 0x7f | (data[pos++] << 7));
			return type;
		}

		int ReadRecordLen()
		{
			int accum = 0;
			int shift = 0;
			for (int i = 0; i < 4; i++, shift += 7)
			{
				if (pos >= end)
				{
					FillBuffer(1);
				}
				var b = data[pos++];
				accum |= (b & 0x7f) << shift;
				if ((b & 0x80) == 0)
					break;
			}
			return accum;
		}

		public bool NextRecord()
		{

			type = ReadRecordType();
			recordLen = ReadRecordLen();
			if(pos + recordLen > end)
			{
				FillBuffer(recordLen);
			}
			s = pos;
			pos += recordLen;

			Debug.Assert(pos <= end);

			return true;
		}
	}
}

enum RecordType
{
	Row = 0,
	CellBlank = 1,
	CellRK = 2,
	CellError = 3,
	CellBool = 4,
	CellNum = 5,
	CellSt = 6,
	CellIsst = 7,
	CellFmlaString = 8,
	CellFmlaNum = 9,
	CellFmlaBool = 10,
	CellFmlaError = 11,
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