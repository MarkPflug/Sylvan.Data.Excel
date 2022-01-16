using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.Common;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Sylvan.Data.Excel
{
	// excel file format specs:
	// https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-xlsb/acc8aa92-1f02-4167-99f5-84f9f676b95a
	// https://docs.microsoft.com/en-us/openspecs/office_file_formats/MS-OFFFFLP/8aea05e3-8c1e-4a9a-9614-31f71e679456
	// https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-oleps/bf7aeae8-c47a-4939-9f45-700158dac3bc

	enum State
	{
		None = 0,
		Initializing,
		Initialized,
		Open,
		End,
		Closed,
	}

	sealed partial class XlsWorkbookReader : ExcelDataReader
	{
		const int Biff8VersionCode = 0x0600;
		const int Biff8EntryDataSize = 8224;
		const int RowBatchSize = 32;

		int rowCount;
		RecordReader reader;
		short biffVersion = 0;

		int yearOffset = 1900;
		int xfIdx = 0;

		Row[] rowBatch = new Row[RowBatchSize];
		CellData[][] rowDatas = new CellData[RowBatchSize][];

		int batchOffset = 0;
		int batchIdx = 0;
		int batchCount = 0;

		bool nullAsEmptyString;
		bool getErrorAsNull;

		int rS = 0;
		int rE = 0;
		int cS = 0;
		int cE = 0;

		int sheetIdx = 0;
		int rowNumber;

		bool closed = false;
		int epoch;

		string[] sst;
		Dictionary<int, ExcelFormat> formats;
		List<SheetInfo> sheets = new List<SheetInfo>();
		Dictionary<int, XFRecord> xfRecords = new Dictionary<int, XFRecord>();

		internal static async Task<XlsWorkbookReader> CreateAsync(Stream iStream, ExcelDataReaderOptions options)
		{
			var pkg = new Ole2Package(iStream);
			var part = pkg.GetEntry("Workbook\0");
			if (part == null)
				throw new InvalidDataException();
			var ps = part.Open();

			var reader = new XlsWorkbookReader(ps, options);
			await reader.ReadHeaderAsync();
			await reader.NextResultAsync();
			return reader;
		}

		private XlsWorkbookReader(Stream iStream, ExcelDataReaderOptions options) : base(options.Schema)
		{
			this.epoch = 1900;
			this.reader = new RecordReader(iStream);
			this.nullAsEmptyString = options.GetNullAsEmptyString;
			this.getErrorAsNull = options.GetErrorAsNull;

			this.columnSchema = new ReadOnlyCollection<DbColumn>(Array.Empty<DbColumn>());
			this.sst = Array.Empty<string>();
			this.formats = ExcelFormat.CreateFormatCollection();
		}

		public override int WorksheetCount => this.sheets.Count;

		public override ExcelWorkbookType WorkbookType => ExcelWorkbookType.Excel;

		public override string WorksheetName
		{
			get
			{
				return
					sheetIdx <= this.sheets.Count
					? this.sheets[sheetIdx - 1].name
					: throw new InvalidOperationException();
			}
		}

		internal override int DateEpochYear => epoch;

		public override ExcelFormat? GetFormat(int ordinal)
		{
			var cell = GetCell(ordinal);
			XFRecord xf = this.xfRecords[cell.ifx];

			if (formats.TryGetValue(xf.ifmt, out var fmt))
			{
				return fmt;
			}
			return null;
		}

		public override ExcelErrorCode GetFormulaError(int ordinal)
		{
			var cell = GetCell(ordinal);
			if (cell.type == CellType.Error)
				return (ExcelErrorCode)cell.val;
			throw new InvalidOperationException();
		}

		public override int RowNumber => rowNumber;

		public override ExcelDataType GetExcelDataType(int ordinal)
		{
			var cell = GetCell(ordinal);
			return cell.type switch
			{
				CellType.Null => ExcelDataType.Null,
				CellType.Boolean => ExcelDataType.Boolean,
				CellType.Error => ExcelDataType.Error,
				CellType.Double => ExcelDataType.Numeric,
				CellType.String => ExcelDataType.String,
				_ => ExcelDataType.Null
			};
		}

		public override bool IsClosed
		{
			get { return this.closed; }
		}


		public override void Close()
		{
			this.closed = true;
		}

		public override async Task<bool> NextResultAsync(CancellationToken cancellationToken)
		{
			if (sheetIdx++ < this.sheets.Count)
			{
				while (Read())
				{
					// process any remaining content in the current sheet
				}

				batchOffset = 0;
				await InitSheet().ConfigureAwait(false);
				return true;
			}
			return false;
		}

		public override bool NextResult()
		{
			return NextResultAsync().GetAwaiter().GetResult();
		}

		public override async Task<bool> ReadAsync(CancellationToken cancellationToken)
		{
			if (this.rowNumber > rowCount)
			{
				return false;
			}
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

			rowNumber++;
			return rowBatch[batchIdx].rowFieldCount > 0;
		}

		public override bool Read()
		{
			return ReadAsync(CancellationToken.None).GetAwaiter().GetResult();
		}

		public override string GetName(int ordinal)
		{
			return columnSchema[ordinal].ColumnName;
		}

		public override Type GetFieldType(int ordinal)
		{
			if (ordinal < 0 || ordinal >= this.columnSchema.Count)
				throw new ArgumentOutOfRangeException(nameof(ordinal));
			return this.columnSchema[ordinal].DataType ?? typeof(string);
		}

		public override int GetOrdinal(string name)
		{
			for (int i = 0; i < this.columnSchema.Count; i++)
			{
				if (string.Compare(this.columnSchema[i].ColumnName, name, false) == 0)
					return i;
			}

			for (int i = 0; i < this.columnSchema.Count; i++)
			{
				if (string.Compare(this.columnSchema[i].ColumnName, name, true) == 0)
					return i;
			}
			throw new ArgumentOutOfRangeException(nameof(name));
		}

		public override int RowFieldCount
		{
			get
			{
				return this.rowBatch[batchIdx].rowFieldCount;
			}
		}

		public override int RowCount => this.rowCount;

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

			while (true)
			{
				await reader.NextRecordAsync();
				var recordType = reader.Type;
				switch (recordType)
				{
					case RecordType.Sst:
						await LoadSharedStringTable();
						break;
					case RecordType.Sheet:
						await LoadSheetRecord();
						break;
					case RecordType.Style:
						ParseStyle();
						break;
					case RecordType.XF:
						ParseXF();
						break;
					case RecordType.Format:
						await ParseFormat();
						break;
					case RecordType.ColInfo:
						//ParseColInfo();
						break;
					case RecordType.EOF:
						return;
					default:
						Debug.WriteLine($"Header: {recordType:x}");
						break;
				}
			}
		}

		async Task<bool> InitSheet()
		{
			rowNumber = 0;
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
						Debug.WriteLine(reader.Type);
						break;
				}
			}
		done:
			await NextRowBatch().ConfigureAwait(false);
			return LoadSchema();
		}


		void ParseXF()
		{
			short ifnt = reader.ReadInt16();
			short ifmt = reader.ReadInt16();
			short flags = reader.ReadInt16();
			XFRecordType type = ((flags & 0x04) == 0) ? XFRecordType.Cell : XFRecordType.Style;

			int parentIdx = (flags & 0xfff0) >> 4;

			xfRecords.Add(xfIdx, new XFRecord { ifmt = ifmt, ifnt = ifnt, ixfParent = parentIdx, type = type });

			Debug.Assert(type == XFRecordType.Cell || parentIdx == 0xfff); // style records must always be 0xfff

			xfIdx++;
		}

		async Task ParseFormat()
		{
			int ifmt = reader.ReadInt16();
			string str;
			if (biffVersion == 0x0500)
				str = await reader.ReadByteString(1);
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
		bool LoadSchema()
		{
			var sheetName = sheets[sheetIdx - 1].name;

			var hasHeaders = schema.HasHeaders(sheetName);

			if (!Read())
			{
				return false;
			}
			LoadSchema(!hasHeaders);

			if (!hasHeaders)
			{
				// "unread" the first row.
				batchIdx--;
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

				double rkVal = RKVal(rk);
				SetRowData(rowIdx, colIdx++, new CellData(rkVal, ixfe));
			}
		}

		async Task ParseLabel()
		{
			int rowIdx = reader.ReadUInt16();
			int colIdx = reader.ReadUInt16();
			int xfIdx = reader.ReadUInt16();
			string str = await reader.ReadByteString(2);
			SetRowData(rowIdx, colIdx, new CellData(str));
		}

		void ParseLabelSST()
		{
			int rowIdx = reader.ReadUInt16();
			int colIdx = reader.ReadUInt16();
			int xfIdx = reader.ReadUInt16();
			int strIdx = reader.ReadInt32();

			SetRowData(rowIdx, colIdx, new CellData(sst[strIdx]));
		}

		void ParseRK()
		{
			int rowIdx = reader.ReadUInt16();
			int colIdx = reader.ReadUInt16();
			ushort xfIdx = reader.ReadUInt16();
			int rk = reader.ReadInt32();

			double rkVal = RKVal(rk);
			SetRowData(rowIdx, colIdx, new CellData(rkVal, xfIdx));
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
			SetRowData(rowIdx, colIdx, new CellData(d, xfIdx));
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
						SetRowData(rowIdx, colIdx, new CellData(str));
						break;
					case 1: // boolean
						SetRowData(rowIdx, colIdx, new CellData(rval, CellType.Boolean));
						break;
					case 2: // error
						SetRowData(rowIdx, colIdx, new CellData(rval, CellType.Error));
						break;
					default:
						throw new InvalidDataException();
				}
			}
			else
			{
				double d = BitConverter.Int64BitsToDouble((long)val);
				SetRowData(rowIdx, colIdx, new CellData(d, xfIdx));
			}
		}

		string FormatVal(int xfIdx, double val)
		{
			XFRecord xf = this.xfRecords[xfIdx];
			var fmtIdx = xf.ifmt;
			if (formats.TryGetValue(fmtIdx, out var fmt))
			{
				return fmt.FormatValue(val, this.yearOffset);
			}
			else
			{
				throw new FormatException();
			}
		}

		static double RKVal(int rk)
		{
			bool mult = (rk & 0x01) != 0;
			bool isFloat = (rk & 0x02) == 0;
			double d;

			if (isFloat)
			{
				long v = rk & 0xfffffffc;
				v = v << 32;
				d = BitConverter.Int64BitsToDouble(v);
			}
			else
			{
				d = rk >> 2;
			}

			if (mult)
			{
				d = d / 100;
			}

			return d;
		}

		void SetRowData(int rowIdx, int colIdx, CellData cd)
		{
			int offset = rowIdx - batchOffset;

			if (offset < 0 || offset >= RowBatchSize)
				throw new IOException(); //cell refers to row that is not in the current batch

			ref var rb = ref rowBatch[offset];
			bool isNull = cd.type == CellType.Null;
			if (!isNull)
			{
				rb.rowFieldCount = Math.Max(rb.firstColIdx, colIdx + 1);
			}
			int rowOff = rb.firstColIdx;
			rowDatas[offset][colIdx - rowOff] = cd;

		}

		async Task<bool> NextRowBatch()
		{
			batchIdx = -1;
			batchCount = 0;
			Array.Clear(this.rowBatch, 0, this.rowBatch.Length);

			do
			{
				await reader.NextRecordAsync();
				switch (reader.Type)
				{
					case RecordType.Row:
						batchCount++;
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

			if (rowDatas[idx] == null || rowDatas[idx].Length < rowLen)
			{
				rowDatas[idx] = new CellData[rowLen];
			}
			for (int i = 0; i < rowLen; i++)
			{
				rowDatas[idx][i] = default;
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

		async Task LoadSheetRecord()
		{
			reader.ReadInt32();
			byte visibility = reader.ReadByte();
			byte type = reader.ReadByte();

			string name =
				biffVersion == 0x0500
				? await reader.ReadByteString(1)
				: await reader.ReadString8();

			sheets.Add(new SheetInfo(type, visibility, name));
		}

		async Task LoadSharedStringTable()
		{
			int totalString = reader.ReadInt32();
			int uniqueString = reader.ReadInt32();

			var strings = new string[uniqueString];

			for (int i = 0; i < uniqueString; i++)
			{
				strings[i] = await reader.ReadString16();
			}

			this.sst = strings;
		}

		public override bool IsDBNull(int ordinal)
		{
			ref var cell = ref GetCell(ordinal);
			switch (cell.type)
			{
				case CellType.Null:
					return true;
				case CellType.Boolean:
				case CellType.Double:
					return false;
				case CellType.Error:
					return
						this.getErrorAsNull
						? true
						: throw new ExcelFormulaException(ordinal, rowNumber, (ExcelErrorCode)cell.val);
				case CellType.String:
				default:
					return cell.str == null;
			}
		}

		ref CellData GetCell(int ordinal)
		{
			ref var r = ref this.rowBatch[batchIdx];

			if (r.index == 0) return ref CellData.Null;

			int rowOffset = r.firstColIdx;
			if (ordinal < rowOffset || ordinal >= r.lastColIdx)
				return ref CellData.Null;

			int dataIdx = ordinal - rowOffset;
			var row = this.rowDatas[batchIdx];
			if (dataIdx < 0 || dataIdx >= row.Length)
				return ref CellData.Null;

			return ref row[dataIdx];
		}


		public override string GetString(int ordinal)
		{
			ref var cell = ref GetCell(ordinal);
			switch (cell.type)
			{
				case CellType.String:
					return cell.str!;
				case CellType.Double:
					return FormatVal(cell.ifx, cell.dVal);
				case CellType.Boolean:
					return cell.val != 0 ? bool.TrueString : bool.FalseString;
				case CellType.Error:
					if (this.getErrorAsNull && this.nullAsEmptyString)
						return string.Empty;
					var errorCode = (ExcelErrorCode)cell.val;
					throw new ExcelFormulaException(ordinal, -1, errorCode);
				case CellType.Null:
					if (this.nullAsEmptyString)
					{
						return string.Empty;
					}
					// GetString is documented to throw this
					throw new InvalidCastException();
			}
			// shouldn't get here.
			throw new NotSupportedException();
		}

		public override double GetDouble(int ordinal)
		{
			ref var cell = ref GetCell(ordinal);
			switch (cell.type)
			{
				case CellType.String:
					return double.Parse(cell.str!);
				case CellType.Double:
					return cell.dVal;
				case CellType.Error:
					throw Error(ordinal);
			}

			throw new FormatException();
		}

		internal override DateTime GetDateTimeValue(int ordinal)
		{
			// only xlsx persists date values this way.
			// in xls files date/time are always stored as formatted numeric values.
			throw new NotSupportedException();
		}

		ExcelFormulaException Error(int ordinal)
		{
			var cell = GetCell(ordinal);
			return new ExcelFormulaException(ordinal, RowNumber, (ExcelErrorCode)cell.val);
		}

		public override bool GetBoolean(int ordinal)
		{
			var cell = GetCell(ordinal);
			switch (cell.type)
			{
				case CellType.Boolean:
					return cell.val != 0;
				case CellType.Double:
					return cell.dVal != 0;
				case CellType.String:
					return bool.Parse(cell.str!);
				case CellType.Error:
					throw new ExcelFormulaException(ordinal, RowNumber, (ExcelErrorCode)cell.val);
				case CellType.Null:
				default:
					throw new InvalidCastException();
			}
		}

		class XFRecord
		{
			public int ifnt;
			public int ifmt;
			public XFRecordType type;
			public int ixfParent;
		}

		struct CellData
		{
			internal static CellData Null = default;

			public CellData(string str)
			{
				this = default;
				this.type = CellType.String;
				this.str = str;
			}

			public CellData(uint val, CellType type)
			{
				this = default;
				this.val = val;
				this.type = type;
			}

			public CellData(double val, ushort ifIdx)
			{
				this = default;
				this.type = CellType.Double;
				this.str = null;
				this.dVal = val;
				this.ifx = ifIdx;
			}

			public string? str;
			public ushort ifx;
			public double dVal;
			public uint val;

			public CellType type;

			public override string ToString()
			{
				switch (type)
				{
					case CellType.Double:
						return "Double: " + dVal;
					case CellType.String:
						return "String: " + str;
				}
				return "NULL";
			}
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
				return $"{index} { firstColIdx} {lastColIdx} {ixfe}";
			}
#endif
		}

		class SheetInfo
		{
			public byte visibility;
			public byte type;
			public string name;

			public SheetInfo(byte type, byte vis, string name)
			{
				this.type = type;
				this.visibility = vis;
				this.name = name;
			}
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

		enum XFRecordType
		{
			Cell = 0,
			Style = 1,
		}

		enum CellType
		{
			Null = 0,
			String,
			Double,
			Boolean,
			Error,
		}
	}
}
