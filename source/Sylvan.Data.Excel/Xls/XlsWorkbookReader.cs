using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Sylvan.Data.Excel;

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

	const int Biff8VersionCode = 0x0600;
	const int Biff8EntryDataSize = 8224;

	RecordReader reader;
	short biffVersion = 0;

	FieldInfo[] fieldInfos;

	int rowIndex;
	int rowNumber = 0;

	int curFieldCount = 0;
	int pendingRow = -1;

	int epoch;

	internal static async Task<XlsWorkbookReader> CreateAsync(Stream iStream, ExcelDataReaderOptions options)
	{
		var reader = new XlsWorkbookReader(iStream, options);
		await reader.ReadHeaderAsync();
		await reader.NextResultAsync();
		return reader;
	}

	private XlsWorkbookReader(Stream stream, ExcelDataReaderOptions options) : base(stream, options)
	{
		var pkg = new Ole2Package(stream);
		var part = pkg.GetEntry("Workbook\0");
		if (part == null)
			throw new InvalidDataException();
		var ps = part.Open();

		this.epoch = 1900;
		this.reader = new RecordReader(ps);
		this.fieldInfos = new FieldInfo[16];
	}

	public override ExcelWorkbookType WorkbookType => ExcelWorkbookType.Excel;

	internal override int DateEpochYear => epoch;

	public override int RowNumber => rowNumber;

	private protected override async Task<bool> OpenWorksheetAsync(int sheetIdx, CancellationToken cancel)
	{
		var info = (XlsSheetInfo)this.sheetInfos[sheetIdx];
		this.rowNumber = 0;
		this.pendingRow = -1;
		reader.SetPosition(info.Offset);
		await InitSheet(cancel).ConfigureAwait(false);
		this.sheetIdx = sheetIdx;
		return true;
	}

	public override async Task<bool> NextResultAsync(CancellationToken cancel)
	{
		sheetIdx++;

		for (; sheetIdx < this.sheetInfos.Length; sheetIdx++)
		{
			if (this.readHiddenSheets || this.sheetInfos[sheetIdx].Hidden == false)
			{
				await OpenWorksheetAsync(sheetIdx, cancel).ConfigureAwait(false);
				return true;
			}
		}
		return false;
	}

	public override bool NextResult()
	{
		return NextResultAsync(default).GetAwaiter().GetResult();
	}

	public override async Task<bool> ReadAsync(CancellationToken cancel)
	{
		rowNumber++;
		if (this.rowIndex >= rowCount)
		{
			rowNumber = -1;
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

		return await NextRow();
	}

	public override bool Read()
	{
		return ReadAsync(CancellationToken.None).GetAwaiter().GetResult();
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

	async Task ReadHeaderAsync()
	{
		await reader.NextRecordAsync();

		if (reader.Type != RecordType.BOF)
			throw new InvalidDataException();//"Expected BOF record"

		BOFType type = ReadBOF();
		if (type != BOFType.WorkbookGlobals)
			throw new InvalidDataException();//"First Stream must be workbook globals stream"
		var sheets = new List<XlsSheetInfo>();
		var xfs = new List<int>();
		bool atEndOfHeader = false;
		while (!atEndOfHeader)
		{
			await reader.NextRecordAsync();
			var recordType = reader.Type;
			switch (recordType)
			{
				case RecordType.Sst:
					await LoadSharedStringTable();
					break;
				case RecordType.Sheet:
					sheets.Add(await LoadSheetRecord());
					break;
				case RecordType.Style:
					ParseStyle();
					break;
				case RecordType.XF:
					xfs.Add(ParseXF());
					break;
				case RecordType.Format:
					await ParseFormat();
					break;
				case RecordType.EOF:
					atEndOfHeader = true;
					break;
				default:
					//Debug.WriteLine($"Header: {recordType:x} {recordType}");
					break;
			}
		}
		this.sheetInfos = sheets.ToArray();
		this.xfMap = xfs.ToArray();
	}

	async Task<bool> InitSheet(CancellationToken cancel)
	{
		rowIndex = -1;
		this.state = State.Initializing;

		while (await reader.NextRecordAsync().ConfigureAwait(false))
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
			await reader.NextRecordAsync().ConfigureAwait(false);

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
		await ReadAsync(cancel).ConfigureAwait(false);
		var result = LoadSchema();
		if (!result)
		{
			await ReadAsync(cancel).ConfigureAwait(false);
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

	async Task ParseFormat()
	{
		int ifmt = reader.ReadInt16();
		string str;
		if (biffVersion == 0x0500)
		{
			str = await reader.ReadByteString(1);
		}
		else
		{
			str = await reader.ReadString16();
		}

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

		if (yearOffsetValue == 1)
		{
			//this.epoch = 1904;
			// don't have the ability to create/test such a file
			// so this will have to remain unsupported.
			// I doubt many such files exist anymore anyway.
			throw new NotSupportedException();
		}
		this.epoch = 1900;
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

	async Task ParseLabel()
	{
		int rowIdx = reader.ReadUInt16();
		int colIdx = reader.ReadUInt16();
		int xfIdx = reader.ReadUInt16();
		string str = await reader.ReadByteString(2);
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

	async Task ParseFormula()
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
					var next = await reader.NextRecordAsync();
					if (!next || reader.Type != RecordType.String) throw new InvalidDataException();
					int len = reader.ReadUInt16();
					byte kind = reader.ReadByte();
					var str = await reader.ReadStringAsync(len, kind == 0);
					SetRowData(colIdx, new FieldInfo(str));
					break;
				case 1: // boolean
					SetRowData(colIdx, new FieldInfo(rval != 0));
					break;
				case 2: // error
					SetRowData(colIdx, new FieldInfo((ExcelErrorCode)rval));
					break;
				default:
					throw new InvalidDataException();
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
		while (colIdx >= fieldInfos.Length)
		{
			Array.Resize(ref fieldInfos, fieldInfos.Length * 2);
		}
		rowFieldCount = Math.Max(rowFieldCount, colIdx + 1);
		fieldInfos[colIdx] = cd;
	}


	async Task<bool> NextRow()
	{
		// clear out any fields from previous row
		Array.Clear(this.fieldInfos, 0, this.fieldInfos.Length);
		this.rowFieldCount = 0;
		do
		{
			if (pendingRow == -1)
			{
				await reader.NextRecordAsync();
			}

			if (rowIndex < pendingRow)
			{
				return true;
			}

			pendingRow = -1;

			switch (reader.Type)
			{
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
							pendingRow = peekRow;
							return true;
						}
						else
						{
							throw new InvalidDataException();
						}
					}
					break;
				case RecordType.EOF:
					if (this.rowFieldCount > 0)
					{
						if (pendingRow == int.MinValue)
						{
							return false;
						}
						else
						{
							pendingRow = int.MinValue;
							return true;
						}
					}
					break;
				default:
					break;
			}

			switch (reader.Type)
			{
				case RecordType.LabelSST:
					ParseLabelSST();
					break;
				case RecordType.Label:
					await ParseLabel();
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
					await ParseFormula();
					break;
				case RecordType.Blank:
				case RecordType.BoolErr:
				case RecordType.MulBlank:
				case RecordType.RString:
					break;
				case RecordType.Array:
				case RecordType.SharedFmla:
				case RecordType.DataTable:
					break;
				case RecordType.String:
					// this should only apply to formulas, and is handled inline
					break;
				case RecordType.EOF:
					return this.RowFieldCount > 0;
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

	async Task<XlsSheetInfo> LoadSheetRecord()
	{
		int offset = reader.ReadInt32();
		byte visibility = reader.ReadByte();
		byte type = reader.ReadByte();

		string name =
			biffVersion == 0x0500
			? await reader.ReadByteString(1)
			: await reader.ReadString8();

		return new XlsSheetInfo(name, offset, visibility != 0);
	}

	async Task LoadSharedStringTable()
	{
		int totalString = reader.ReadInt32();
		int uniqueString = reader.ReadInt32();

		var strings = new string[uniqueString];

		for (int i = 0; i < uniqueString; i++)
		{
			var s = await reader.ReadString16();
			strings[i] = s;
		}

		this.sst = strings;
	}

	private protected override ref readonly FieldInfo GetFieldValue(int ordinal)
	{
		if (ordinal >= this.fieldInfos.Length)
			return ref FieldInfo.Null;

		return ref this.fieldInfos[ordinal];
	}

	internal override DateTime GetDateTimeValue(int ordinal)
	{
		// only xlsx persists date values this way.
		// in xls files date/time are always stored as formatted numeric values.
		throw new NotSupportedException();
	}
}
