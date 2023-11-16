#nullable enable
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace Sylvan.Data.Excel.Xlsb;

sealed class XlsbWorkbookReader : ExcelDataReader
{
	readonly ZipArchive package;

	Stream sheetStream;
	RecordReader? reader;

	bool hasRows = false;
	//bool skipEmptyRows = true; // TODO: make this an option?

	int rowIndex;
	int parsedRowIndex = -1;
	int curFieldCount = -1;

	readonly ZipArchiveEntry? sstPart;
	Stream? sstStream;
	RecordReader? sstReader;
	int sstIdx = -1;

	public override ExcelWorkbookType WorkbookType => ExcelWorkbookType.ExcelXml;

	public override void Close()
	{
		this.sheetStream?.Close();
		this.sstStream?.Close();
		base.Close();
	}

	const string DefaultWorkbookPartName = "xl/workbook.bin";

	public XlsbWorkbookReader(Stream stream, ExcelDataReaderOptions opts) : base(stream, opts)
	{
		this.sheetStream = Stream.Null;
		package = new ZipArchive(stream, ZipArchiveMode.Read);

		var workbookPartName = OpenPackaging.GetWorkbookPart(package) ?? DefaultWorkbookPartName;

		var workbookPart = package.GetEntry(workbookPartName);

		var sheetsRelsPart = OpenPackaging.GetPartRelationsName(workbookPartName);

		var stylesPartName = "xl/styles.bin";
		var sharedStringsPartName = "xl/sharedStrings.bin";

		if (workbookPart == null || sheetsRelsPart == null)
			throw new InvalidDataException();

		var sheetRelMap = OpenPackaging.LoadWorkbookRelations(package, workbookPartName, ref stylesPartName, ref sharedStringsPartName);

		var stylePart = package.GetEntry(stylesPartName);

		this.sstPart = package.GetEntry(sharedStringsPartName);

		var sheetNameList = new List<SheetInfo>();
		using (Stream sheetsStream = workbookPart.Open())
		{
			var rr = new RecordReader(sheetsStream);
			while (rr.NextRecord())
			{
				switch (rr.RecordType)
				{
					case RecordType.WbProp:
						var f = rr.GetInt32();
						this.dateMode = ((f & 1) == 1)
							? DateMode.Mode1904
							: DateMode.Mode1900;
						break;

					case RecordType.BundleSheet:
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
						break;
				}
				//rr.DebugInfo("sheets");
			}
		}

		this.sheetInfos = sheetNameList.ToArray();
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

	private protected override ref readonly FieldInfo GetFieldValue(int ordinal)
	{
		if (rowIndex < parsedRowIndex || ordinal >= this.RowFieldCount)
			return ref FieldInfo.Null;

		return ref values[ordinal];
	}

	private protected override bool OpenWorksheet(int sheetIdx)
	{
		var sheetName = sheetInfos[sheetIdx].Part;
		// the relationship is recorded as an absolute path
		// but the zip entry has a relative name.
		sheetName = sheetName.TrimStart('/');

		var sheetPart = package.GetEntry(sheetName);
		if (sheetPart == null)
			throw new InvalidDataException();

		if (sheetStream != null)
		{
			this.sheetStream.Close();
		}
		this.sheetStream = sheetPart.Open();

		this.reader = new RecordReader(this.sheetStream);
		var rr = this.reader;
		this.rowFieldCount = 0;
		this.curFieldCount = -1;
		this.sheetIdx = sheetIdx;
		return InitializeSheet();
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
			return false;

		return OpenWorksheet(sheetIdx);
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
			if (!reader.NextRecord())
				throw new InvalidDataException();

			if (reader.RecordType == RecordType.Dimension)
			{
				var rowLast = reader.GetInt32(4);
				var colLast = reader.GetInt32(12);
				if (this.values.Length < colLast + 1)
				{
					Array.Resize(ref values, colLast + 1);
				}
				this.rowCount = rowLast + 1;
			}
			if (reader.RecordType == RecordType.DataStart)
			{
				break;
			}
		}

		var c = ParseRowValues();

		if (c == -1)
		{
			return false;
		}

		if (parsedRowIndex > 0)
		{
			this.curFieldCount = this.rowFieldCount;
			this.rowFieldCount = 0;
		}

		if (LoadSchema())
		{
			this.state = State.Initialized;
			this.rowIndex = -1;
		}
		else
		{
			this.state = State.Open;
			this.rowIndex = 0;
		}

		return true;
	}

	bool LoadSst(int idx)
	{
		var reader = this.sstReader;
		if (sstPart == null)
		{
			return false;
		}
		if (reader == null)
		{
			this.sstStream = sstPart.Open();
			reader = this.sstReader = new RecordReader(this.sstStream);
			reader.NextRecord();
			if (reader.RecordType != RecordType.SSTBegin)
				throw new InvalidDataException();

			int totalCount = reader.GetInt32(0);
			int count = reader.GetInt32(4);

			if (count > 128)
				count = 128;
			this.sst = new string[count];
		}
		while (idx > this.sstIdx)
		{
			if (!reader.NextRecord() || reader.RecordType != RecordType.SSTItem)
			{
				throw new InvalidDataException();
			}

			var flags = reader.GetByte(0);
			var str = reader.GetString(1);
			this.sstIdx++;
			if (sstIdx >= this.sst.Length)
			{
				Array.Resize(ref sst, sst.Length * 2);
			}
			sst[sstIdx] = str;
		}
		return true;
	}

	private protected override string GetSharedString(int idx)
	{
		if (this.sstIdx < idx)
		{
			if (!LoadSst(idx))
			{
				throw new InvalidDataException();
			}
		}
		return sst[idx];
	}

	public override bool Read()
	{
		rowIndex++;

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

			while (true)
			{
				var c = ParseRowValues();
				if (c < 0)
				{
					this.rowFieldCount = 0;
					return false;
				}
				if (c == 0)
				{
					continue;
				}
				if (rowIndex < parsedRowIndex)
				{
					this.curFieldCount = c;
					this.rowFieldCount = 0;
					return true;
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
				if (rowIndex == parsedRowIndex && curFieldCount >= 0)
				{
					this.rowFieldCount = curFieldCount;
					this.curFieldCount = -1;
				}
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

			static void EnsureCols(ref FieldInfo[] values, int c)
			{
				if (values.Length <= c)
					Array.Resize(ref values, c + 8);
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
						EnsureCols(ref values, col);
						ref var fi = ref values[col];

						fi.type = ExcelDataType.Numeric;
						fi.numValue = d;
						fi.xfIdx = sf;
						notNull++;
						count = col + 1;
					}
					break;
				case RecordType.CellReal:
					{
						var col = reader.GetInt32(0);
						var sf = reader.GetInt32(4) & 0xffffff;
						double d = reader.GetDouble(8);
						EnsureCols(ref values, col);
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

						EnsureCols(ref values, col);
						ref var fi = ref values[col];

						switch (reader.RecordType)
						{
							case RecordType.CellBlank:
								type = ExcelDataType.Null;
								break;
							case RecordType.CellBool:
							case RecordType.CellFmlaBool:
								type = ExcelDataType.Boolean;
								fi = new FieldInfo(reader.GetByte(8) != 0);
								notNull++;
								break;
							case RecordType.CellError:
							case RecordType.CellFmlaError:
								type = ExcelDataType.Error;
								fi = new FieldInfo((ExcelErrorCode)reader.GetByte(8));
								notNull++;
								break;
							case RecordType.CellIsst:
								type = ExcelDataType.String;
								var sstIdx = reader.GetInt32(8);
								
								fi.isSS = true;
								fi.ssIdx = sstIdx;
								//fi.strValue = sst[sstIdx];
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

	internal override DateTime GetDateTimeValue(int ordinal)
	{
		throw new NotSupportedException();
	}

	public override int MaxFieldCount => 16384;

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

		bool FillBuffer(int requiredLen)
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
					return false;

				end += l;
			}
			Debug.Assert(pos <= end);
			return true;
		}

		RecordType ReadRecordType()
		{
			Debug.Assert(pos <= end);

			if (pos >= end)
			{
				if (!FillBuffer(1))
				{
					return RecordType.None;
				}
			}

			var b = data[pos++];
			if ((b & 0x80) == 0)
			{
				return (RecordType)b;
			}

			if (pos >= end)
			{
				if (!FillBuffer(1))
				{
					// the second byte wasn't there.
					throw new InvalidDataException();
				}
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
			if (type == RecordType.None)
			{
				return false;
			}
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
}
