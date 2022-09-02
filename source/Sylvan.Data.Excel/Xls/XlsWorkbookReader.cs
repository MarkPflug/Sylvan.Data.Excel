#nullable enable

using System;
using System.Collections.Generic;
using System.Diagnostics;
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
	const int RowBatchSize = 32;

	RecordReader reader;
	short biffVersion = 0;

	Row[] rowBatch = new Row[RowBatchSize];
	FieldInfo[][] fieldInfos = new FieldInfo[RowBatchSize][];

	int batchOffset = 0;
	int batchIdx = 0;

	int rS = 0;
	int rE = 0;
	int cS = 0;
	int cE = 0;

	int rowIndex;
	int parsedRowIndex;

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
	}

	public override ExcelWorkbookType WorkbookType => ExcelWorkbookType.Excel;

	internal override int DateEpochYear => epoch;

	public override int RowNumber => rowIndex + 1;

	public override async Task<bool> NextResultAsync(CancellationToken cancellationToken)
	{
		sheetIdx++;
		for (; sheetIdx < this.sheetNames.Length; sheetIdx++)
		{
			var info = (XlsSheetInfo) this.sheetNames[sheetIdx];

			reader.SetPosition(info.Offset);

			batchOffset = 0;
			await InitSheet().ConfigureAwait(false);
			if (this.readHiddenSheets || this.sheetNames[sheetIdx].Hidden == false)
			{
				return true;
			}
		}
		return false;
	}

	public override bool NextResult()
	{
		return NextResultAsync(default).GetAwaiter().GetResult();
	}

	public override async Task<bool> ReadAsync(CancellationToken cancellationToken)
	{
		if (this.rowIndex >= rowCount)
		{
			return false;
		}
		if (state == State.Initialized)
		{
			this.state = State.Open;
			return true;
		}
		rowIndex++;

		// "catch up" to the next non-empty row
		if (rowIndex <= parsedRowIndex)
		{
			return true;
		}

		// look for a row that has values.
		// this is needed to trim the tail of rows that are empty.
		while (parsedRowIndex < rowCount)
		{
			parsedRowIndex++;
			batchIdx++;

			if (batchIdx >= RowBatchSize)
			{
				if (await NextRowBatch())
				{
					batchIdx = 0;
				}
				else
				{
					return false;
				}
			}

			this.rowFieldCount = this.rowBatch[batchIdx].rowFieldCount;

			if (rowBatch[batchIdx].rowFieldCount > 0)
			{
				// found a row with values
				return true;
			}
		}
		return false;
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
				case RecordType.ColInfo:
					//ParseColInfo();
					break;
				case RecordType.EOF:
					atEndOfHeader = true;
					break;
				default:
					Debug.WriteLine($"Header: {recordType:x} {recordType}");
					break;
			}
		}
		this.sheetNames = sheets.ToArray();
		this.xfMap = xfs.ToArray();
	}

	async Task<bool> InitSheet()
	{
		rowIndex = -1;
		parsedRowIndex = -1;

		while (await reader.NextRecordAsync().ConfigureAwait(false))
		{
			if (reader.Type == RecordType.BOF)
			{
				BOFType type = ReadBOF();
				switch (type)
				{
					case BOFType.Worksheet:
					case BOFType.Biff4MacroSheet:
						goto go;
					case BOFType.Chart:
						continue;
					default:
						throw new NotSupportedException();
				}
				throw new InvalidDataException();//"Expected sheetBOF"
			}
		}
		throw new InvalidDataException();//"Expected sheetBOF"
	go:

		while (true)
		{
			await reader.NextRecordAsync().ConfigureAwait(false);

			switch (reader.Type)
			{
				case RecordType.ColInfo:
					//ParseColInfo();
					break;
				case RecordType.Dimension:
					ParseDimension();
					this.rowCount = rE;
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
		await NextRowBatch().ConfigureAwait(false);
		return LoadSchemaXls();
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

	// return value indicates if there are any rows in the sheet.
	bool LoadSchemaXls()
	{
		var sheetName = this.WorksheetName;
		if (sheetName == null) return false;
		if (!Read())
		{
			return false;
		}
		if (LoadSchema())
		{
			// "unread" the first row.
			rowIndex--;
		}

		return true;
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
			SetRowData(rowIdx, colIdx++, new FieldInfo(rkVal, ixfe));
		}
	}

	async Task ParseLabel()
	{
		int rowIdx = reader.ReadUInt16();
		int colIdx = reader.ReadUInt16();
		int xfIdx = reader.ReadUInt16();
		string str = await reader.ReadByteString(2);
		SetRowData(rowIdx, colIdx, new FieldInfo(str));
	}

	void ParseLabelSST()
	{
		int rowIdx = reader.ReadUInt16();
		int colIdx = reader.ReadUInt16();
		int xfIdx = reader.ReadUInt16();
		int strIdx = reader.ReadInt32();

		SetRowData(rowIdx, colIdx, new FieldInfo(sst[strIdx]));
	}

	void ParseRK()
	{
		int rowIdx = reader.ReadUInt16();
		int colIdx = reader.ReadUInt16();
		ushort xfIdx = reader.ReadUInt16();
		int rk = reader.ReadInt32();

		double rkVal = GetRKVal(rk);
		SetRowData(rowIdx, colIdx, new FieldInfo(rkVal, xfIdx));
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
		SetRowData(rowIdx, colIdx, new FieldInfo(d, xfIdx));
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

		//var opts = reader.ReadUInt16();
		//var chn = reader.ReadInt32();
		//var flen = reader.ReadUInt16();

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
					SetRowData(rowIdx, colIdx, new FieldInfo(str));
					break;
				case 1: // boolean
					SetRowData(rowIdx, colIdx, new FieldInfo(rval != 0));
					break;
				case 2: // error
					SetRowData(rowIdx, colIdx, new FieldInfo((ExcelErrorCode)rval));
					break;
				default:
					throw new InvalidDataException();
			}
		}
		else
		{
			double d = BitConverter.Int64BitsToDouble((long)val);
			SetRowData(rowIdx, colIdx, new FieldInfo(d, xfIdx));
		}
	}

	void SetRowData(int rowIdx, int colIdx, FieldInfo cd)
	{
		int offset = rowIdx - batchOffset;

		if (offset < 0 || offset >= RowBatchSize)
			throw new IOException(); //cell refers to row that is not in the current batch

		ref var rb = ref rowBatch[offset];
		bool isNull = cd.type == ExcelDataType.Null;
		if (!isNull)
		{
			rb.rowFieldCount = Math.Max(rb.firstColIdx, colIdx + 1);
		}
		int rowOff = rb.firstColIdx;
		fieldInfos[offset][colIdx - rowOff] = cd;
	}

	async Task<bool> NextRowBatch()
	{
		batchIdx = -1;
		Array.Clear(this.rowBatch, 0, this.rowBatch.Length);

		do
		{
			await reader.NextRecordAsync();
			switch (reader.Type)
			{
				case RecordType.Row:
					ParseRow();
					break;
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
				case RecordType.Blank:
				case RecordType.BoolErr:
				case RecordType.MulBlank:
				case RecordType.RString:
					break;
				case RecordType.Formula:
					await ParseFormula();
					break;
				case RecordType.Array:
				case RecordType.SharedFmla:
				case RecordType.DataTable:
					break;
				case RecordType.String:
					// this should only apply to formulas, and is handled inline
					break;
				case RecordType.DBCell:
					batchOffset += RowBatchSize;
					return true;
				case RecordType.EOF:
					batchOffset += RowBatchSize;
					return false;
				default:
					break;
			}
		} while (true);
	}

	void ParseRow()
	{
		var rowIdx = reader.ReadUInt16();
		var firstColIdx = reader.ReadUInt16();
		var lastColIdx = reader.ReadUInt16();

		int rowHeight = reader.ReadUInt16();
		int reserved = reader.ReadUInt16();
		int reserved2 = reader.ReadUInt16();

		var flags = reader.ReadUInt16();

		var ixfe = reader.ReadUInt16();

		if ((flags & 0x80) != 0)
			ixfe = (ushort)(ixfe & 0x0fff);
		else
			ixfe = 0x0fff;

		var idx = rowIdx - batchOffset;

		ref var r = ref rowBatch[idx];

		r.index = rowIdx + 1;
		r.firstColIdx = firstColIdx;
		r.lastColIdx = lastColIdx;
		r.ixfe = ixfe;

		int rowLen = lastColIdx - firstColIdx;

		if (fieldInfos[idx] == null || fieldInfos[idx].Length < rowLen)
		{
			fieldInfos[idx] = new FieldInfo[rowLen];
		}
		for (int i = 0; i < rowLen; i++)
		{
			fieldInfos[idx][i] = default;
		}
	}

	void ParseDimension()
	{
		if (biffVersion == 0x0500)
		{
			rS = reader.ReadUInt16();
			rE = reader.ReadUInt16();
		}
		else
		{
			rS = reader.ReadInt32();
			rE = reader.ReadInt32();
		}
		cS = reader.ReadUInt16();
		cE = reader.ReadUInt16();

		reader.ReadUInt16();

		if (rS > rE || cS > cE)
			throw new InvalidDataException();
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
		if (rowIndex < parsedRowIndex)
			return ref FieldInfo.Null;

		ref var r = ref this.rowBatch[batchIdx];

		if (r.index == 0) return ref FieldInfo.Null;

		int rowOffset = r.firstColIdx;
		if (ordinal < rowOffset || ordinal >= r.lastColIdx)
			return ref FieldInfo.Null;

		int dataIdx = ordinal - rowOffset;
		var row = this.fieldInfos[batchIdx];
		if (dataIdx < 0 || dataIdx >= row.Length)
			return ref FieldInfo.Null;

		return ref row[dataIdx];
	}

	internal override DateTime GetDateTimeValue(int ordinal)
	{
		// only xlsx persists date values this way.
		// in xls files date/time are always stored as formatted numeric values.
		throw new NotSupportedException();
	}

	struct Row
	{
		public int rowFieldCount;
		public int index;
		public ushort firstColIdx;
		public ushort lastColIdx;
		public ushort ixfe;

#if DEBUG
		public override string ToString()
		{
			return $"{index} {firstColIdx} {lastColIdx} {ixfe}";
		}
#endif
	}

	enum RecordType
	{
		Dimension = 0x0200,
		YearEpoch = 0x022,
		Blank = 0x0201,
		Number = 0x0203,
		Label = 0x0204,
		BoolErr = 0x0205,
		Formula = 0x0006,
		String = 0x0207,

		BOF = 0x0809,

		Continue = 0x003c,
		CRN = 0x005a,
		LabelSST = 0x00fd,

		RK = 0x027e,

		MulRK = 0x00BD,
		EOF = 0x000A,
		XF = 0x00e0,

		Font = 0x0031,
		ExtSst = 0x00ff,
		Format = 0x041e,
		Style = 0x0293,
		Row = 0x0208,

		ExternSheet = 0x0017,
		DefinedName = 0x0018,
		Country = 0x008c,

		Index = 0x020B,

		CalcCount = 0x000c,
		CalcMode = 0x000d,
		Precision = 0x000e,
		RefMode = 0x000f,

		Delta = 0x0010,
		Iteration = 0x0011,
		Protect = 0x0012,
		Password = 0x0013,
		Header = 0x0014,
		Footer = 0x0015,
		ExternCount = 0x0016,

		Guts = 0x0080,
		SheetPr = 0x0081,
		GridSet = 0x0082,
		HCenter = 0x0083,
		VCenter = 0x0084,
		Sheet = 0x0085,
		WriteProt = 0x0086,

		Sort = 0x0090,

		ColInfo = 0x007d,

		Sst = 0x00fc,
		MulBlank = 0x00be,
		RString = 0x00d6,
		Array = 0x0221,
		SharedFmla = 0x04bc,
		DataTable = 0x0236,
		DBCell = 0x00d7,
	}

	enum BOFType
	{
		WorkbookGlobals = 0x0005,
		VisualBasicModule = 0x0006,
		Worksheet = 0x0010,
		Chart = 0x0020,
		Biff4MacroSheet = 0x0040,
		Biff4WorkbookGlobals = 0x0100,
	}
}
