#nullable enable
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace Sylvan.Data.Excel;

sealed class XlsbWorkbookReader : ExcelDataReader
{
	static readonly Dictionary<int, ExcelFormat> EmptyFormats = new Dictionary<int, ExcelFormat>(0);

	const string RelationsNS = "http://schemas.openxmlformats.org/package/2006/relationships";

	Dictionary<int, ExcelFormat> formats = EmptyFormats;
	int[] xfMap = Array.Empty<int>();

	string[] stringData;

	readonly ZipArchive package;
	int sheetIdx = -1;
	int rowCount;

	Stream sheetStream;
	RecordReader? reader;

	FieldInfo[] values;
	int rowFieldCount;
	State state;
	bool hasRows = false;
	bool skipEmptyRows = true; // TODO: make this an option?

	int rowIndex;
	int parsedRowIndex;

	SheetInfo[] sheetNames;

	bool readHiddenSheets;
	bool errorAsNull;

	struct FieldInfo
	{
		public ExcelDataType type;
		public string strValue;
		public double numValue;
		public int xfIdx;
		public ExcelErrorCode err;
		public bool b;
	}

	public override int RowCount => rowCount;

	public override ExcelWorkbookType WorkbookType => ExcelWorkbookType.ExcelXml;

	class SheetInfo
	{
		public SheetInfo(string name, string part, bool hidden)
		{
			this.Name = name;
			this.Part = part;
			this.Hidden = hidden;
		}

		public string Name { get; }
		public string Part { get; }
		public bool Hidden { get; }
	}

	public override void Close()
	{
		this.sheetStream?.Close();
		base.Close();
	}

	public XlsbWorkbookReader(Stream stream, ExcelDataReaderOptions opts) : base(stream, opts.Schema)
	{
		this.rowCount = -1;
		this.values = Array.Empty<FieldInfo>();
		this.errorAsNull = opts.GetErrorAsNull;
		this.readHiddenSheets = opts.ReadHiddenWorksheets;

		this.sheetStream = Stream.Null;
		package = new ZipArchive(stream, ZipArchiveMode.Read);

		var stylePart = package.GetEntry("xl/styles.bin");

		var sheetsPart = package.GetEntry("xl/workbook.bin");
		var sheetsRelsPart = package.GetEntry("xl/_rels/workbook.bin.rels");

		if (sheetsPart == null)
			throw new InvalidDataException();

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

		stringData = ReadSharedStrings();

		var sheetNameList = new List<SheetInfo>();
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
								var hs = rr.GetInt32(0);
								var hidden = hs != 0;
								var id = rr.GetInt32(4);
								var rel = rr.GetString(8, out int next);
								var name = rr.GetString(next);
								if (rel == null)
								{
									// no sheet rel means it is a macro.
								}
								else
								{
									var part = sheetRelMap[rel!];
									var info = new SheetInfo(name, part, hidden);
									sheetNameList.Add(info);
								}
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

		this.sheetNames = sheetNameList.ToArray();
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
		
	public override bool NextResult()
	{
		sheetIdx++;
		for (; sheetIdx < this.sheetNames.Length; sheetIdx++)
		{
			if (readHiddenSheets || sheetNames[sheetIdx].Hidden == false)
			{
				break;
			}
		}
		if (sheetIdx >= this.sheetNames.Length)
			return false;

		var sheetName = sheetNames[sheetIdx].Part;
		// the relationship is recorded as an absolute path
		// but the zip entry has a relative name.
		sheetName = sheetName.TrimStart('/');

		var sheetPart = package.GetEntry(sheetName);
		if (sheetPart == null)
			return false;
		if (sheetStream != null)
		{
			this.sheetStream.Close();
		}
		this.sheetStream = sheetPart.Open();

		this.reader = new RecordReader(this.sheetStream);
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

		var hasHeaders = schema.HasHeaders(this.WorksheetName!);

		LoadSchema(!hasHeaders);

		if (hasHeaders)
		{
			this.state = State.Open;
			Read();
		}

		this.rowIndex = hasHeaders ? 0 : -1;
		this.state = State.Initialized;
		return true;
	}

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
				//if (flags == 0)
				//{
				var str = reader.GetString(1);
				ss[i] = str;
				//}
				//else
				//{
				//	throw new NotImplementedException();
				//}
			}
			return ss;
		}
	}

	public override bool Read()
	{
		rowIndex++;

		if (state == State.Open)
		{
			if (rowIndex <= parsedRowIndex)
			{
				return true;
			}

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
		rowIndex = -1;
		this.state = State.End;
		return false;
	}

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

						var d = GetRKVal(rk);

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

		this.parsedRowIndex = rowIdx;
		this.rowFieldCount = count;
		return notNull;
	}

	public override string GetName(int ordinal)
	{
		return this.columnSchema?[ordinal].ColumnName ?? "";
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
		throw new NotSupportedException();
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
					throw GetError(ordinal);
			}
		}
		throw new InvalidCastException();
	}

	public override string GetString(int ordinal)
	{
		if (rowIndex < parsedRowIndex)
		{
			return string.Empty;
		}
		if (ordinal >= MaxFieldCount)
			throw new ArgumentOutOfRangeException(nameof(ordinal));
		if (ordinal >= rowFieldCount)
			return String.Empty;
		ref var fi = ref values[ordinal];
		switch (fi.type)
		{
			case ExcelDataType.Error:
				if (errorAsNull)
				{
					return string.Empty;
				}
				throw GetError(ordinal);
			case ExcelDataType.Boolean:
				return fi.b ? bool.TrueString : bool.FalseString;
			case ExcelDataType.Numeric:
				return FormatVal(fi.xfIdx, fi.numValue);
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

	public override int MaxFieldCount => 16384;

	public override int WorksheetCount => this.sheetNames.Length;

	public override string? WorksheetName => sheetIdx < sheetNames.Length ? this.sheetNames[this.sheetIdx].Name : null;

	internal override int DateEpochYear => 1900;

	public override int RowNumber => rowIndex + 1;

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

		public string? GetString(int offset, out int end)
		{
			var len = BitConverter.ToInt32(data, s + offset);
			if (len == -1)
			{
				end = offset + 4;
				return null;
			}
			end = offset + 4 + len * 2;
			return Encoding.Unicode.GetString(data, s + offset + 4, len * 2);
		}

		void FillBuffer(int requiredLen)
		{
			Debug.Assert(pos <= end);

			if (this.data.Length < requiredLen)
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

			if (pos >= end)
			{
				FillBuffer(1);
				if (pos >= end)
					throw new EndOfStreamException();
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
			if (pos + recordLen > end)
			{
				FillBuffer(recordLen);
			}
			s = pos;
			pos += recordLen;

			Debug.Assert(pos <= end);

			return true;
		}
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
}

