﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;

namespace Sylvan.Data.Excel.Xls;

sealed partial class XlsWorkbookReader : ExcelDataReader
{
	sealed class XlsSheetInfo : SheetInfo
	{
		public XlsSheetInfo(string name, int offset, bool hidden) : base(name, hidden)
		{
			this.Offset = offset;
		}

		public int Offset { get; }
	}

	// rows are stored in blocks of records
	const int RowBatchSize = 32;
	const int Biff8VersionCode = 0x0600;
	const int Biff8EntryDataSize = 8224;

	RecordReader reader;
	short biffVersion = 0;

	int rowNumber = 0;

	int curFieldCount = 0;
	int pendingRow = -1;

	int rowCellCount = 0;

	internal XlsWorkbookReader(Stream stream, ExcelDataReaderOptions options) : base(stream, options)
	{
		var pkg = new Ole2Package(stream);
		var part =
			pkg.GetEntry("Workbook\0") ??
			pkg.GetEntry("Book\0");

		if (part == null)
			throw new InvalidDataException();
		var ps = part.Open();

		this.reader = new RecordReader(ps);
		this.ReadHeader();
		this.NextResult();
	}

	BitArray rowHidden = new BitArray(32);

	public override ExcelWorkbookType WorkbookType => ExcelWorkbookType.Excel;

	public override int RowNumber => rowNumber;

	public override bool IsRowHidden
	{
		get
		{
			return this.rowHidden[(this.rowNumber - 1) & 0b11111];
		}
	}

	private protected override bool OpenWorksheet(int sheetIdx)
	{
		var info = (XlsSheetInfo)this.sheetInfos[sheetIdx];
		this.rowNumber = 0;
		this.pendingRow = -1;
		reader.SetPosition(info.Offset);
		InitSheet();
		this.sheetIdx = sheetIdx;
		return true;
	}

	public override bool NextResult()
	{
		sheetIdx++;

		for (; sheetIdx < this.sheetInfos.Length; sheetIdx++)
		{
			if (this.readHiddenSheets || this.sheetInfos[sheetIdx].Hidden == false)
			{
				OpenWorksheet(sheetIdx);
				return true;
			}
		}
		return false;
	}

	public override bool Read()
	{
	next:
		rowNumber++;
		colCacheIdx = 0;
		if (this.rowIndex >= rowCount)
		{
			rowNumber = -1;
			this.state = State.End;
			return false;
		}
		if (state == State.Initialized)
		{
			this.state = State.Open;
			this.rowFieldCount = this.curFieldCount;
			this.curFieldCount = 0;
			return true;
		}

		rowIndex++;

		var count = NextRow();

		if (count < 0)
		{
			if (this.rowCellCount > 0 && this.ignoreEmptyTrailingRows == false)
			{
				return true;
			}
			this.state = State.End;
			return false;
		}
		else
		{
			if (this.readHiddenRows == false && this.IsRowHidden)
			{
				goto next;
			}
			return true;
		}
	}

	private protected override string GetSharedString(int idx)
	{
		// .xls eagerly loads the shared strings.
		return sst[idx];
	}

	public override int MaxFieldCount => 256;

	BOFType ReadBOF()
	{
		short ver = reader.ReadInt16();
		if (biffVersion == 0)
			biffVersion = ver;
		if (!(biffVersion == 0x0600 || biffVersion == 0x0500))
			throw new InvalidDataException();//"Invalid stream version"

		short type = reader.ReadInt16();
		return (BOFType)type;
	}

	void ReadHeader()
	{
		reader.NextRecord();

		if (reader.Type != RecordType.BOF)
			throw new InvalidDataException();//"Expected BOF record"

		BOFType type = ReadBOF();
		if (type != BOFType.WorkbookGlobals)
			throw new InvalidDataException();//"First Stream must be workbook globals stream"
		var sheets = new List<XlsSheetInfo>();
		var xfs = new List<int>();

		while (reader.NextRecord())
		{
			var recordType = reader.Type;
			switch (recordType)
			{
				case RecordType.Sst:
					LoadSharedStringTable();
					break;
				case RecordType.Sheet:
					sheets.Add(LoadSheetRecord());
					break;
				case RecordType.Style:
					ParseStyle();
					break;
				case RecordType.XF:
					xfs.Add(ParseXF());
					break;
				case RecordType.Format:
					ParseFormat();
					break;
				case RecordType.YearEpoch:
					Parse1904();
					break;
				case RecordType.EOF:
					goto done;
				default:
					//Debug.WriteLine($"Header: {recordType:x} {recordType}");
					break;
			}
		}
	done:
		this.sheetInfos = sheets.ToArray();
		this.xfMap = xfs.ToArray();
	}

	bool InitSheet()
	{
		rowIndex = -1;
		this.state = State.Initializing;

		while (reader.NextRecord())
		{
			if (reader.Type == RecordType.BOF)
			{
				BOFType type = ReadBOF();
				switch (type)
				{
					case BOFType.Worksheet:
					case BOFType.Biff4MacroSheet:
						goto readBeginningOfSheet;
					case BOFType.Chart:
						continue;
					default:
						throw new NotSupportedException();
				}
				throw new InvalidDataException();//"Expected sheetBOF"
			}
		}
		throw new InvalidDataException();//"Expected sheetBOF"

	readBeginningOfSheet:
		while (true)
		{
			reader.NextRecord();

			switch (reader.Type)
			{
				case RecordType.ColInfo:
					//ParseColInfo();
					break;
				case RecordType.Dimension:
					this.rowCount = ParseDimension();
					goto done;
				case RecordType.YearEpoch:
					Parse1904();
					break;
				case RecordType.EOF:
					throw new InvalidDataException();//"Unexpected EOF"
				default:
					//Debug.WriteLine(reader.Type);
					break;
			}
		}
	done:
		Read();
		var result = LoadSchema();
		if (!result)
		{
			Read();
			this.rowNumber = 1;
		}
		else
		{
			this.rowNumber = 0;
		}
		this.curFieldCount = this.rowFieldCount;
		this.rowFieldCount = this.FieldCount;
		this.state = State.Initialized;
		return result;
	}

	int ParseXF()
	{
		short ifnt = reader.ReadInt16();
		short ifmt = reader.ReadInt16();
		short flags = reader.ReadInt16();

		return ifmt;
	}

	void ParseFormat()
	{
		int ifmt = reader.ReadInt16();
		string str =
			biffVersion == 0x0500
			? reader.ReadByteString(1)
			: reader.ReadString16();

		if (formats.ContainsKey(ifmt))
		{
			formats.Remove(ifmt);
		}

		var fmt = new ExcelFormat(str);
		formats.Add(ifmt, fmt);
	}

	void ParseStyle()
	{
		// ignoring styles, at least for now.
	}

	void Parse1904()
	{
		int yearOffsetValue = reader.ReadInt16();

		this.dateMode =
			yearOffsetValue == 1
			? DateMode.Mode1904
			: DateMode.Mode1900;
	}

	void ParseMulRK()
	{
		int count = (reader.Length - 6) / 6;

		int rowIdx = reader.ReadUInt16();
		int colIdx = reader.ReadUInt16();

		for (int i = 0; i < count; i++)
		{
			var ixfe = reader.ReadUInt16();
			int rk = reader.ReadInt32();

			double rkVal = GetRKVal(rk);
			SetRowData(colIdx++, new FieldInfo(rkVal, ixfe));
		}
	}

	void ParseLabel()
	{
		int rowIdx = reader.ReadUInt16();
		int colIdx = reader.ReadUInt16();
		int xfIdx = reader.ReadUInt16();
		int len = reader.ReadInt16();
		if (len > 255) throw new InvalidDataException();
		bool compressed = true;
		if (biffVersion == 0x0500)
		{
			// apparently there are no flags in this version
		}
		else
		{
			byte flags = reader.ReadByte();
			compressed = (flags & 1) == 0;
		}

		var str = reader.ReadStringBuffer(len, compressed);
		SetRowData(colIdx, new FieldInfo(str));
	}

	void ParseRString()
	{
		int rowIdx = reader.ReadUInt16();
		int colIdx = reader.ReadUInt16();
		int xfIdx = reader.ReadUInt16();
		var len = reader.ReadInt16();
		var str = reader.ReadStringBuffer(len, true);

		// consume the formatting info
		var x = reader.ReadByte();
		for (int i = 0; i < x; i++)
		{
			reader.ReadUInt16();
		}

		SetRowData(colIdx, new FieldInfo(str));
	}

	void ParseLabelSST()
	{
		int rowIdx = reader.ReadUInt16();
		int colIdx = reader.ReadUInt16();
		int xfIdx = reader.ReadUInt16();
		int strIdx = reader.ReadInt32();

		SetRowData(colIdx, new FieldInfo(sst[strIdx]));
	}

	void ParseRK()
	{
		int rowIdx = reader.ReadUInt16();
		int colIdx = reader.ReadUInt16();
		ushort xfIdx = reader.ReadUInt16();
		int rk = reader.ReadInt32();

		double rkVal = GetRKVal(rk);
		SetRowData(colIdx, new FieldInfo(rkVal, xfIdx));
	}

	void ParseNumber()
	{
		int rowIdx = reader.ReadUInt16();
		int colIdx = reader.ReadUInt16();
		ushort xfIdx = reader.ReadUInt16();
		long val;
		unchecked
		{
			uint uL = (uint)reader.ReadInt32();
			uint uH = (uint)reader.ReadInt32();
			val = ((long)uL) | ((long)uH << 32);
		}
		double d = BitConverter.Int64BitsToDouble(val);
		SetRowData(colIdx, new FieldInfo(d, xfIdx));
	}

	void ParseFormula()
	{
		var rowIdx = reader.ReadUInt16();
		var colIdx = reader.ReadUInt16();
		var xfIdx = reader.ReadUInt16();

		ulong val;
		unchecked
		{
			uint uL = (uint)reader.ReadInt32();
			uint uH = (uint)reader.ReadInt32();
			val = ((ulong)uL) | ((ulong)uH << 32);
		}

		// if the 2 MSB of the value are 0xff, then the stored value
		// is not a number, but is a string, boolean, or error
		if ((val & 0xffff_0000_0000_0000ul) == 0xffff_0000_0000_0000ul)
		{
			var rtype = (int)(val & 0xff);
			var rval = (uint)(val >> 16 & 0xff);
			switch (rtype)
			{
				case 0: // string
					var next = reader.NextRecord();
					if (!next || reader.Type != RecordType.String) throw new InvalidDataException();
					int len = reader.ReadUInt16();
					byte kind = reader.ReadByte();
					var str = reader.ReadString(len, kind == 0);
					SetRowData(colIdx, new FieldInfo(str));
					break;
				case 1: // boolean
					SetRowData(colIdx, new FieldInfo(rval != 0));
					break;
				case 2: // error
					SetRowData(colIdx, new FieldInfo((ExcelErrorCode)rval));
					break;
				default:
					// this seems to indicate the function result is null,
					// though the spec doesn't make this clear.
					SetRowData(colIdx, new FieldInfo());
					break;
			}
		}
		else
		{
			double d = BitConverter.Int64BitsToDouble((long)val);
			SetRowData(colIdx, new FieldInfo(d, xfIdx));
		}
	}

	void SetRowData(int colIdx, FieldInfo cd)
	{
		if (colIdx >= MaxFieldCount)
			throw new InvalidDataException();
		// TODO: this could be cleaner
		while (colIdx >= values.Length)
		{
			Array.Resize(ref values, Math.Max(8, values.Length * 2));
		}
		if (!cd.IsEmptyValue)
		{
			this.rowFieldCount = Math.Max(rowFieldCount, colIdx + 1);
		}
		this.rowCellCount++;
		values[colIdx] = cd;
	}


	int NextRow()
	{
		// clear out any fields from previous row
		Array.Clear(this.values, 0, this.values.Length);
		// rowFieldCount records the last non-empty cell.
		this.rowFieldCount = 0;
		// rowCellCount records the number of cells that have any (even empty string) values
		this.rowCellCount = 0;

		do
		{
			if (pendingRow == -1)
			{
				if (!reader.NextRecord())
				{
					// reached the end of the records stream before finding any more cells
					return -1;
				}
			}

			if (rowIndex < pendingRow)
			{
				// the current row is empty but there is more data after.
				return 0;
			}

			pendingRow = -1;

			// this first switch is only concerned with "peeking" at the next cell record
			// to determine if it is for the current row (rowIndex), or if the current row
			// is empty where the next cell is for a subsequent row.
			switch (reader.Type)
			{
				//case RecordType.Row:
				case RecordType.LabelSST:
				case RecordType.Label:
				case RecordType.RK:
				case RecordType.MulRK:
				case RecordType.Number:
				case RecordType.Formula:
					// inspect the row of the next cell without advancing the reader
					var peekRow = reader.PeekRow();
					if (this.rowIndex != peekRow)
					{
						if (this.rowIndex < peekRow)
						{
							// the current row is empty but we've seen a cell for a subsequent row.
							pendingRow = peekRow;
							return 0;
						}
						else
						{
							peekRow = (ushort)(rowIndex + 1);
							pendingRow = peekRow;
							return 0;
							//throw new InvalidDataException();
						}
					}
					break;
				case RecordType.EOF:
					if (this.rowFieldCount > 0)
					{
						// we've reached the end of the data stream
						// and have cells in the current row
						if (pendingRow == int.MinValue)
						{
							return -1;
						}
						else
						{
							// set pending row such that we will come back to return -1
							// the next time we read a row.
							pendingRow = int.MinValue;
							return 0;
						}
					}
					break;
				default:
					break;
			}

			switch (reader.Type)
			{
				case RecordType.Row:
					// TODO: I should really be handling this more similarly to .xlsb reading.
					// where I can read from a specific offset in the record.
					var r1 = reader.ReadInt16();
					var r2 = reader.ReadInt16();
					reader.ReadInt32();
					reader.ReadInt32();
					var flags = reader.ReadInt32();
					this.rowHidden[r1 & 0b11111] = (flags & 0x20) != 0;
					break;
				case RecordType.LabelSST:
					ParseLabelSST();
					break;
				case RecordType.Label:
					ParseLabel();
					break;
				case RecordType.RK:
					ParseRK();
					break;
				case RecordType.MulRK:
					ParseMulRK();
					break;
				case RecordType.Number:
					ParseNumber();
					break;
				case RecordType.Formula:
					ParseFormula();
					break;
				case RecordType.RString:
					ParseRString();
					break;
				case RecordType.Blank:
				case RecordType.BoolErr:
				case RecordType.MulBlank:
					break;
				case RecordType.Array:
				case RecordType.SharedFmla:
				case RecordType.DataTable:
					break;
				case RecordType.String:
					// this should only apply to formulas, and is handled inline
					break;
				case RecordType.EOF:
					return this.rowFieldCount == 0 ? -1 : this.rowFieldCount;
				default:
					break;
			}
		} while (true);
	}

	int ParseDimension()
	{
		int rowStart, rowEnd;
		if (biffVersion == 0x0500)
		{
			rowStart = reader.ReadUInt16();
			rowEnd = reader.ReadUInt16();
		}
		else
		{
			rowStart = reader.ReadInt32();
			rowEnd = reader.ReadInt32();
		}
		var colStart = reader.ReadUInt16();
		var colEnd = reader.ReadUInt16();

		reader.ReadUInt16();

		if (rowStart > rowEnd || colStart > colEnd)
			throw new InvalidDataException();

		return rowEnd;
	}

	XlsSheetInfo LoadSheetRecord()
	{
		int offset = reader.ReadInt32();
		byte visibility = reader.ReadByte();
		byte type = reader.ReadByte();

		string name =
			biffVersion == 0x0500
			? reader.ReadByteString(1)
			: reader.ReadString8();

		return new XlsSheetInfo(name, offset, visibility != 0);
	}

	void LoadSharedStringTable()
	{
		int totalString = reader.ReadInt32();
		int uniqueString = reader.ReadInt32();

		var strings = new string[uniqueString];

		for (int i = 0; i < uniqueString; i++)
		{
			var s = reader.ReadString16();
			strings[i] = s;
		}

		this.sst = strings;
	}

	private protected override ref readonly FieldInfo GetFieldValue(int ordinal)
	{
		if (ordinal >= this.values.Length)
			return ref FieldInfo.Null;

		return ref this.values[ordinal];
	}

	internal override DateTime GetDateTimeValue(int ordinal)
	{
		// only xlsx persists date values this way.
		// in xls files date/time are always stored as formatted numeric values.
		throw new NotSupportedException();
	}
}
