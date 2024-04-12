#nullable enable
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.Common;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace Sylvan.Data.Excel;

/// <summary>
/// A DbDataReader implementation that reads data from an Excel file.
/// </summary>
public abstract partial class ExcelDataReader : DbDataReader, IDisposable, IDbColumnSchemaGenerator
{
	// Excel supports two different date representations.
	internal enum DateMode
	{
		Mode1900,
		Mode1904,
	}

	int fieldCount;
	bool isClosed;
	Stream stream;
#pragma warning disable
	bool isAsync; // currently unused, but intend to use it to enforce async access patterns.
#pragma warning restore
	private protected IExcelSchemaProvider schema;
	private protected State state;
	private protected ExcelColumn[] columnSchema;
	private protected bool ownsStream;
	private protected Dictionary<int, ExcelFormat> formats;
	private protected int[] xfMap;
	private protected FieldInfo[] values;
	private protected string[] sst;

	private protected SheetInfo[] sheetInfos;
	private protected int sheetIdx = -1;

	private protected bool readHiddenSheets;
	private protected bool errorAsNull;

	private protected int rowCount;
	private protected int rowFieldCount;


	private protected int rowIndex;

	static readonly DateTime Epoch1900 = new DateTime(1899, 12, 30);
	static readonly DateTime Epoch1904 = new DateTime(1904, 1, 1);

	private protected DateMode dateMode;

	readonly string? trueString;
	readonly string? falseString;

	readonly CultureInfo culture;
	readonly string? dateTimeFormat;

	/// <inheritdoc/>
	public sealed override Type GetFieldType(int ordinal)
	{
		ValidateSheetRange(ordinal);
		if (ordinal < fieldCount)
		{
			return this.columnSchema[ordinal].DataType ?? typeof(object);
		}
		return typeof(object);
	}

	/// <inheritdoc/>
	public sealed override DataTable GetSchemaTable()
	{
		return SchemaTable.GetSchemaTable(this.GetColumnSchema());
	}

	private protected ExcelDataReader(Stream stream, ExcelDataReaderOptions options)
	{
		this.isAsync = false;
		this.stream = stream;
		this.schema = options.Schema;
		this.errorAsNull = options.GetErrorAsNull;
		this.readHiddenSheets = options.ReadHiddenWorksheets;
		this.state = State.Initializing;
		this.values = Array.Empty<FieldInfo>();
		this.sst = Array.Empty<string>();

		this.xfMap = Array.Empty<int>();
		this.sheetInfos = Array.Empty<SheetInfo>();

		this.columnSchema = Array.Empty<ExcelColumn>();
		this.formats = ExcelFormat.CreateFormatCollection();

		this.trueString = options.TrueString;
		this.falseString = options.FalseString;
		this.culture = options.Culture;
		this.dateTimeFormat = options.DateTimeFormat;
		this.ownsStream = options.OwnsStream;
	}

#if ASYNC

	/// <summary>
	/// Asynchronously creates a new ExcelDataReader.
	/// </summary>
	/// <param name="filename">The name of the file to open.</param>
	/// <param name="options">An optional ExcelDataReaderOptions instance.</param>
	/// <param name="cancel">A CancellationToken.</param>
	/// <returns>The ExcelDataReader.</returns>
	/// <exception cref="ArgumentException">If the filename refers to a file of an unknown type.</exception>
	public static async Task<ExcelDataReader> CreateAsync(string filename, ExcelDataReaderOptions? options = null, CancellationToken cancel = default)
	{
		var type = GetWorkbookType(filename);
		if (type == ExcelWorkbookType.Unknown)
			throw new ArgumentException(null, nameof(filename));

		var s = File.OpenRead(filename);
		try
		{
			return await CreateAsync(s, type, options, cancel).ConfigureAwait(false);
		}
		finally
		{
			if (s != null)
			{
				var t = s.DisposeAsync();
				await t.AsTask().ConfigureAwait(false);
			}
		}
	}

	/// <summary>
	/// Creates a new ExcelDataReader instance.
	/// </summary>
	/// <param name="stream">A stream containing the Excel file contents. </param>
	/// <param name="fileType">The type of file represented by the stream.</param>
	/// <param name="options">An optional ExcelDataReaderOptions instance.</param>
	/// <param name="cancel"></param>
	/// <returns>The ExcelDataReader.</returns>
	public static async Task<ExcelDataReader> CreateAsync(Stream stream, ExcelWorkbookType fileType, ExcelDataReaderOptions? options = null, CancellationToken cancel = default)
	{
		options ??= ExcelDataReaderOptions.Default;

		var ms = new Sylvan.IO.PooledMemoryStream();
		await stream.CopyToAsync(ms, cancel).ConfigureAwait(false);
		ms.Seek(0, SeekOrigin.Begin);
		try
		{
			ExcelDataReader reader;
			switch (fileType)
			{
				case ExcelWorkbookType.Excel:
					reader = new Xls.XlsWorkbookReader(ms, options);
					break;
				case ExcelWorkbookType.ExcelXml:
					reader = new XlsxWorkbookReader(ms, options);
					break;
				case ExcelWorkbookType.ExcelBinary:
					reader = new Xlsb.XlsbWorkbookReader(ms, options);
					break;
				default:
					throw new ArgumentException(nameof(fileType));
			}
			// In async mode, the reader always owns the memory stream.
			// This causes disposal to dispose the memorystream, and return any pooled buffers.
			reader.ownsStream = true;
			reader.isAsync = true;
			return reader;
		}
		catch
		{
			ms?.Dispose();
			throw;
		}
		finally
		{
			if (options.OwnsStream)
			{
				var t = stream.DisposeAsync();
				await t.AsTask().ConfigureAwait(false);
			}
		}
	}

#endif

	/// <summary>
	/// Creates a new ExcelDataReader.
	/// </summary>
	/// <param name="filename">The name of the file to open.</param>
	/// <param name="options">An optional ExcelDataReaderOptions instance.</param>
	/// <returns>The ExcelDataReader.</returns>
	/// <exception cref="ArgumentException">If the filename refers to a file of an unknown type.</exception>
	public static ExcelDataReader Create(string filename, ExcelDataReaderOptions? options = null)
	{
		var type = GetWorkbookType(filename);
		if (type == ExcelWorkbookType.Unknown)
			throw new ArgumentException(null, nameof(filename));

		var s = File.OpenRead(filename);
		try
		{
			var reader = Create(s, type, options);
			reader.ownsStream = true;
			return reader;
		}
		catch (Exception)
		{
			s?.Dispose();
			throw;
		}
	}

	/// <summary>
	/// Creates a new ExcelDataReader instance.
	/// </summary>
	/// <param name="stream">A stream containing the Excel file contents. </param>
	/// <param name="fileType">The type of file represented by the stream.</param>
	/// <param name="options">An optional ExcelDataReaderOptions instance.</param>
	/// <returns>The ExcelDataReader.</returns>
	public static ExcelDataReader Create(Stream stream, ExcelWorkbookType fileType, ExcelDataReaderOptions? options = null)
	{
		options = options ?? ExcelDataReaderOptions.Default;

		switch (fileType)
		{
			case ExcelWorkbookType.Excel:
				return new Xls.XlsWorkbookReader(stream, options);
			case ExcelWorkbookType.ExcelXml:
				return new XlsxWorkbookReader(stream, options);
			case ExcelWorkbookType.ExcelBinary:
				return new Xlsb.XlsbWorkbookReader(stream, options);
			default:
				throw new ArgumentException(nameof(fileType));
		}
	}

	/// <inheritdoc/>
	public override bool IsClosed => isClosed;

	/// <inheritdoc/>
	public override void Close()
	{
		this.isClosed = true;
		if (ownsStream)
		{
			stream.Dispose();
		}
	}

	/// <summary>
	/// Gets the number of fields in the current row.
	/// This may be different than FieldCount.
	/// </summary>
	public int RowFieldCount => this.rowFieldCount;

	/// <summary>
	/// Gets the maximum number of fields supported by the
	/// file.
	/// </summary>
	public abstract int MaxFieldCount { get; }

	void ValidateSheetRange(int ordinal)
	{
		if ((uint)ordinal >= this.MaxFieldCount)
		{
			throw new ArgumentOutOfRangeException(nameof(ordinal));
		}
	}

	void ValidateAccess()
	{
		if (state != State.Open)
			throw new InvalidOperationException();
	}

	/// <summary>
	/// Gets the type of an Excel workbook from the file name.
	/// </summary>
	public static ExcelWorkbookType GetWorkbookType(string filename)
	{
		return ExcelFileType.FindForFilename(filename)?.WorkbookType ?? ExcelWorkbookType.Unknown;
	}

	/// <summary>
	/// Tries to open a worksheet.
	/// </summary>
	/// <param name="name">The name of the worksheet to open.</param>
	/// <returns>True if the sheet was opened, otherwise false.</returns>
	public bool TryOpenWorksheet(string name)
	{
		var sheetIdx = -1;
		for (int i = 0; i < this.sheetInfos.Length; i++)
		{
			if (StringComparer.OrdinalIgnoreCase.Equals(name, this.sheetInfos[i].Name))
			{
				sheetIdx = i;
				break;
			}
		}
		if (sheetIdx == -1)
		{
			return false;
		}
		return OpenWorksheet(sheetIdx);
	}

	private protected abstract bool OpenWorksheet(int sheetIdx);

	/// <summary>
	/// Gets the names of the worksheets in the workbook.
	/// </summary>
	public IEnumerable<string> WorksheetNames
	{
		get
		{
			return
				this.sheetInfos
				.Where(s => this.readHiddenSheets || !s.Hidden)
				.Select(s => s.Name);
		}
	}

	/// <summary>
	/// Gets the number of worksheets in the workbook.
	/// </summary>
	public int WorksheetCount => this.sheetInfos.Length;

	/// <summary>
	/// Gets the name of the current worksheet.
	/// </summary>
	public string? WorksheetName
	{
		get
		{
			return
				sheetIdx < this.sheetInfos.Length
				? this.sheetInfos[sheetIdx].Name
				: null;
		}
	}

	/// <summary>
	/// Gets the type of workbook being read.
	/// </summary>
	public abstract ExcelWorkbookType WorkbookType { get; }

	/// <summary>
	/// Gets the number of rows in the current sheet.
	/// </summary>
	/// <remarks>
	/// Can return -1 to indicate that the number of rows is unknown.
	/// </remarks>
	public int RowCount => rowCount;

	/// <inheritdoc/>
	public sealed override int FieldCount => this.fieldCount;

	/// <inheritdoc/>
	public sealed override string GetName(int ordinal)
	{
		var cs = this.columnSchema;
		if (cs != null)
		{
			if (ordinal < cs.Length)
			{
				return cs[ordinal].ColumnName;
			}
		}
		return string.Empty;
	}

	/// <inheritdoc/>
	public sealed override int GetOrdinal(string name)
	{
		for (int i = 0; i < this.columnSchema.Length; i++)
		{
			if (string.Compare(this.columnSchema[i].ColumnName, name, false) == 0)
				return i;
		}

		for (int i = 0; i < this.columnSchema.Length; i++)
		{
			if (string.Compare(this.columnSchema[i].ColumnName, name, true) == 0)
				return i;
		}
		throw new ArgumentOutOfRangeException(nameof(name));
	}

	/// <summary>
	/// Gets the type of data in the given cell.
	/// </summary>
	/// <remarks>
	/// Excel only explicitly supports storing either string or numeric (double) values.
	/// Date and Time values are represented by formatting applied to numeric values.
	/// Formulas can produce string, numeric, boolean or error values. 
	/// Boolean and error values are only produced as formula results.
	/// The Null type represents missing rows or cells.
	/// </remarks>
	/// <param name="ordinal">The zero-based column ordinal.</param>
	/// <returns>An ExcelDataType.</returns>
	public ExcelDataType GetExcelDataType(int ordinal)
	{
		ValidateAccess();
		ValidateSheetRange(ordinal);
		ref readonly var cell = ref GetFieldValue(ordinal);
		return cell.type;
	}

	/// <summary>
	/// Gets the value as represented in Excel.
	/// </summary>
	/// <remarks>
	/// Formula errors are returned as ExcelErrorCode values, rather than throwing an exception.
	/// </remarks>
	/// <param name="ordinal">The column ordinal to retrieve.</param>
	/// <returns>The value.</returns>
	public object GetExcelValue(int ordinal)
	{
		ValidateAccess();
		var type = GetExcelDataType(ordinal);
		switch (type)
		{
			case ExcelDataType.Boolean:
				return GetBoolean(ordinal);
			case ExcelDataType.DateTime:
				return GetDateTime(ordinal);
			case ExcelDataType.Error:
				// TODO: cache the boxed values?
				return GetFormulaError(ordinal);
			case ExcelDataType.Null:
				return DBNull.Value;
			case ExcelDataType.Numeric:
				return GetDouble(ordinal);
			case ExcelDataType.String:
				return GetString(ordinal);
			default:
				throw new NotSupportedException();
		}
	}

	/// <summary>
	/// Gets the column schema of the current worksheet.
	/// </summary>
	public ReadOnlyCollection<DbColumn> GetColumnSchema()
	{
		return new ReadOnlyCollection<DbColumn>(this.columnSchema);
	}

	/// <summary>
	/// Initializes the schema starting with the current row.
	/// </summary>
	/// <remarks>
	/// This can be used when a worksheet has "header" rows with non-data content.
	/// Read past the header, and call Initialize when the row of tabular data is found.
	/// </remarks>
	public void Initialize()
	{
		var sheet = this.WorksheetName;
		if (sheet == null)
		{
			throw new InvalidOperationException();
		}

		if (this.state == State.Initialized)
		{
			// prevent reinitializing on the first row, which is already implicitly initialized
			// the values array will already hold the second row of data, so reinitialization
			// isn't possible.
			return;
		}

		if (LoadSchema())
		{
			this.state = State.Initialized;
		}
	}

	private protected bool LoadSchema()
	{
		var sheet = this.WorksheetName;
		if (sheet == null)
			throw new InvalidOperationException();

		var hasHeaders = schema.HasHeaders(sheet);
		var fieldCount = schema.GetFieldCount(this);
		var cols = new ExcelColumn[fieldCount];
		for (int i = 0; i < fieldCount; i++)
		{
			string? header = hasHeaders ? GetStringRaw(i) : null;
			var col = schema.GetColumn(sheet, header, i);
			var ecs = new ExcelColumn(header, i, col);
			cols[i] = ecs;
		}
		this.columnSchema = cols;
		this.fieldCount = fieldCount;

		// return value indicates if the current data row is already
		// sitting in the values array, indicating that the next call
		// to Read() should used the existing values.
		return !hasHeaders;
	}

	/// <inheritdoc/>
	public sealed override int GetValues(object[] values)
	{
		ValidateAccess();
		var c = Math.Min(values.Length, this.FieldCount);

		for (int i = 0; i < c; i++)
		{
			values[i] = this.GetValue(i);
		}
		return c;
	}



	/// <inheritdoc/>
	public sealed override object GetValue(int ordinal)
	{
		ValidateAccess();
		ValidateSheetRange(ordinal);
		if (IsDBNull(ordinal))
			return DBNull.Value;

		var schemaType = GetFieldType(ordinal);
		var code = Type.GetTypeCode(schemaType);

		switch (code)
		{
			case TypeCode.Boolean:
				return GetBoolean(ordinal);
			case TypeCode.Int16:
				return GetInt16(ordinal);
			case TypeCode.Int32:
				return GetInt32(ordinal);
			case TypeCode.Int64:
				return GetInt64(ordinal);
			case TypeCode.Single:
				return GetFloat(ordinal);
			case TypeCode.Double:
				return GetDouble(ordinal);
			case TypeCode.Decimal:
				return GetDecimal(ordinal);
			case TypeCode.DateTime:
				return GetDateTime(ordinal);
			case TypeCode.String:
				return GetString(ordinal);
			default:
				if (schemaType == typeof(Guid))
				{
					return GetGuid(ordinal);
				}
				if (schemaType == typeof(object))
				{
					// when the type of the column is object
					// we'll treat it "dynamically" and try
					// to return the most appropriate value.
					var type = GetExcelDataType(ordinal);
					switch (type)
					{
						case ExcelDataType.Boolean:
							return GetBoolean(ordinal);
						case ExcelDataType.DateTime:
							return GetDateTime(ordinal);
						case ExcelDataType.Error:
							throw GetError(ordinal);
						case ExcelDataType.Null:
							return DBNull.Value;
						case ExcelDataType.Numeric:
							var fmt = GetFormat(ordinal);
							var kind = fmt?.Kind ?? FormatKind.Number;
							switch (kind)
							{
								case FormatKind.Number:
									var doubleValue = GetDouble(ordinal);
									unchecked
									{
										// Excel stores all values as double
										// but we'll try to return it as the
										// most "intuitive" type.
										var int32Value = (int)doubleValue;
										if (doubleValue == int32Value)
											return int32Value;
										var int64Value = (long)doubleValue;
										if (doubleValue == int64Value)
											return int64Value;
										return doubleValue;
									}
								case FormatKind.Date:
									return GetDateTime(ordinal);
								case FormatKind.Time:
									return GetFieldValue<TimeSpan>(ordinal);
							}
							break;
						case ExcelDataType.String:
							return GetString(ordinal);
						default:
							throw new NotSupportedException();
					}
				}
				break;
		}
		throw new NotSupportedException();
	}

	/// <inheritdoc/>
	public sealed override IEnumerator GetEnumerator()
	{
		while (this.Read())
		{
			yield return this;
		}
	}

	/// <inheritdoc/>
	public sealed override int Depth => 0;

	/// <inheritdoc/>
	public sealed override object this[int ordinal] => this.GetValue(ordinal);

	/// <inheritdoc/>
	public sealed override object this[string name] => this.GetValue(this.GetOrdinal(name));

	/// <inheritdoc/>
	public sealed override string GetDataTypeName(int ordinal)
	{
		return this.GetFieldType(ordinal).Name;
	}

	/// <inheritdoc/>
	public sealed override int RecordsAffected => 0;

	/// <inheritdoc/>
	public sealed override bool HasRows => this.RowCount != 0;

	/// <summary>
	/// Gets the <see cref="ExcelErrorCode"/> of the error in the given cell.
	/// </summary>
	public ExcelErrorCode GetFormulaError(int ordinal)
	{
		ValidateAccess();
		var cell = GetFieldValue(ordinal);
		if (cell.type == ExcelDataType.Error)
			return cell.ErrorCode;
		throw new InvalidOperationException();
	}

	internal ExcelFormulaException GetError(int ordinal)
	{
		return new ExcelFormulaException(ordinal, RowNumber, GetFormulaError(ordinal));
	}

	/// <summary>
	/// Gets the <see cref="ExcelFormat"/> of the format for the given cell.
	/// </summary>
	public ExcelFormat? GetFormat(int ordinal)
	{
		ValidateAccess();
		var fi = GetFieldValue(ordinal);
		var idx = fi.xfIdx;

		idx = idx <= 0 ? 0 : xfMap[idx];
		if (this.formats.TryGetValue(idx, out var fmt))
		{
			return fmt;
		}
		return null;
	}

	private protected virtual ref readonly FieldInfo GetFieldValue(int ordinal)
	{
		if (ordinal >= this.rowFieldCount)
			return ref FieldInfo.Null;

		return ref values[ordinal];
	}

	/// <summary>
	/// Gets the number of the current row, as would be reported in Excel.
	/// </summary>
	public abstract int RowNumber { get; }

	/// <summary>
	/// Gets the value of the column as a DateTime.
	/// </summary>
	/// <remarks>
	/// When called on cells containing a string value, will attempt to parse the string as a DateTime.
	/// When called on a cell containing a number value, will convert the numeric value to a DateTime.
	/// </remarks>
	public override DateTime GetDateTime(int ordinal)
	{
		ValidateAccess();
		var type = this.GetExcelDataType(ordinal);
		DateTime value;
		switch (type)
		{
			case ExcelDataType.Boolean:
			case ExcelDataType.Null:
				throw new InvalidCastException();
			case ExcelDataType.Error:
				throw new ExcelFormulaException(ordinal, this.RowNumber, GetFormulaError(ordinal));
			case ExcelDataType.Numeric:
				var val = GetDouble(ordinal);
				var fmt = GetFormat(ordinal) ?? ExcelFormat.Default;
				return TryGetDate(fmt, val, out value)
					? value
					: throw new InvalidCastException();
			case ExcelDataType.DateTime:
				return GetDateTimeValue(ordinal);
			case ExcelDataType.String:
			default:
				var str = GetString(ordinal);
				var fmtStr = this.columnSchema[ordinal]?.Format ?? this.dateTimeFormat;
				if (fmtStr != null)
				{
					if (DateTime.TryParseExact(str, fmtStr, culture, DateTimeStyles.None, out value))
					{
						return value;
					}
				}

				return
					DateTime.TryParse(str, culture, DateTimeStyles.None, out value)
					? value
					: throw new InvalidCastException();
		}
	}

	/// <summary>
	/// Gets the value of the column as a TimeSpan.
	/// </summary>
	/// <remarks>
	/// When called on cells containing a string value, will attempt to parse the string as a TimeSpan.
	/// When called on a cell containing a number value, will convert the numeric value to a DateTime and return the Time component.
	/// </remarks>
	public TimeSpan GetTimeSpan(int ordinal)
	{
		ValidateAccess();
		var type = this.GetExcelDataType(ordinal);
		switch (type)
		{
			case ExcelDataType.Error:
				throw new ExcelFormulaException(ordinal, this.RowNumber, GetFormulaError(ordinal));
			case ExcelDataType.Numeric:
				var val = GetDouble(ordinal);
				var fmt = GetFormat(ordinal);
				if (fmt?.Kind == FormatKind.Time)
				{
					if (TryGetDate(fmt, val, out DateTime dt))
					{
						return dt.TimeOfDay;
					}
				}
				break;
			case ExcelDataType.String:
			default:
				var str = GetString(ordinal);
				if (TimeSpan.TryParse(str, out TimeSpan value))
				{
					return value;
				}
				break;
		}
		throw new InvalidCastException();
	}

	internal abstract DateTime GetDateTimeValue(int ordinal);


	internal bool TryGetDate(ExcelFormat fmt, double value, out DateTime dt)
	{
		return TryGetDate(fmt, value, this.dateMode, out dt);
	}

	static internal bool TryGetDate(ExcelFormat fmt, double value, DateMode mode, out DateTime dt)
	{
		dt = DateTime.MinValue;
		DateTime epoch = Epoch1904;
		// Excel doesn't render negative values as dates.
		if (value < 0.0)
			return false;
		if (mode == DateMode.Mode1900)
		{
			epoch = Epoch1900;
			if (value < 61d)
			{
				if (value < 1)
				{
					if (fmt.Kind == FormatKind.Time)
					{
						dt = DateTime.MinValue.AddDays(value);
						return true;
					}

					// 0 is rendered as 1900-1-0, which is nonsense.
					// negative values render as "###"
					// so we won't support accessing such values.
					return false;
				}
				if (value >= 60d)
				{
					// 1900 wasn't a leapyear, but Excel thinks it was
					// values in this range are in-expressible as .NET dates
					// Excel renders it as 1900-2-29 (not a real day)
					return false;
				}
				value += 1;
			}
		}
		dt = epoch.AddDays(value);
		return true;
	}

	/// <inheritdoc/>
	public sealed override bool IsDBNull(int ordinal)
	{
		ValidateAccess();
		if (ordinal < this.columnSchema.Length && this.columnSchema[ordinal].AllowDBNull == false)
		{
			return false;
		}

		var type = this.GetExcelDataType(ordinal);
		switch (type)
		{
			case ExcelDataType.String:
				return string.IsNullOrEmpty(this.GetString(ordinal));
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



	/// <summary>
	/// Gets the value of the column as a string.
	/// </summary>
	/// <remarks>
	/// With the default configuration, this method is safe to call on all cells.
	/// For cells with missing/null data or a formula error, it will produce an empty string.
	/// </remarks>
	/// <param name="ordinal">The zero-based column ordinal.</param>
	/// <returns>A string representing the value of the column.</returns>
	public sealed override string GetString(int ordinal)
	{
		ValidateAccess();
		return GetStringRaw(ordinal);
	}

	string GetStringRaw(int ordinal) {
		ref readonly FieldInfo fi = ref GetFieldValue(ordinal);
		if (ordinal >= MaxFieldCount)
		{
			throw new ArgumentOutOfRangeException(nameof(ordinal));
		}

		switch (fi.type)
		{
			case ExcelDataType.Error:
				if (errorAsNull)
				{
					return string.Empty;
				}
				throw GetError(ordinal);
			case ExcelDataType.Boolean:
				return fi.BoolValue ? bool.TrueString : bool.FalseString;
			case ExcelDataType.Numeric:
				return FormatVal(fi.xfIdx, fi.numValue);
		}
		return ProcString(in fi);
	}

	string ProcString(in FieldInfo fi)
	{
		return (fi.isSS ? GetSharedString(fi.ssIdx) : fi.strValue) ?? string.Empty;
	}

	private protected abstract string GetSharedString(int idx);

	string FormatVal(int xfIdx, double val)
	{
		var fmtIdx = xfIdx >= this.xfMap.Length ? -1 : this.xfMap[xfIdx];
		if (fmtIdx == -1)
		{
			return val.ToString();
		}

		if (formats.TryGetValue(fmtIdx, out var fmt))
		{
			return fmt.FormatValue(val, this.dateMode);
		}
		else
		{
			return val.ToString();
		}
	}

	/// <inheritdoc/>
	public sealed override float GetFloat(int ordinal)
	{
		ValidateAccess();
		return (float)GetDouble(ordinal);
	}

	/// <inheritdoc/>
	public sealed override double GetDouble(int ordinal)
	{
		ValidateAccess();
		ref readonly var cell = ref GetFieldValue(ordinal);
		switch (cell.type)
		{
			case ExcelDataType.String:
				return double.Parse(ProcString(in cell), culture);
			case ExcelDataType.Numeric:
				return cell.numValue;
			case ExcelDataType.Error:
				throw Error(ordinal);
		}

		throw new InvalidCastException();
	}

	ExcelFormulaException Error(int ordinal)
	{
		ref readonly var cell = ref GetFieldValue(ordinal);
		return new ExcelFormulaException(ordinal, RowNumber, cell.ErrorCode);
	}

	/// <inheritdoc/>
	public sealed override bool GetBoolean(int ordinal)
	{
		ValidateAccess();
		ref readonly var fi = ref this.GetFieldValue(ordinal);
		switch (fi.type)
		{
			case ExcelDataType.Boolean:
				return fi.BoolValue;
			case ExcelDataType.Numeric:
				return this.GetDouble(ordinal) != 0;
			case ExcelDataType.String:

				var col = (uint)ordinal < this.columnSchema.Length ? this.columnSchema[ordinal] : null;

				var trueString = col?.TrueString ?? this.trueString;
				var falseString = col?.FalseString ?? this.falseString;

				var strVal = ProcString(in fi);
				var c = StringComparer.OrdinalIgnoreCase;

				if (trueString != null && c.Equals(strVal, trueString))
				{
					return true;
				}
				if (falseString != null && c.Equals(strVal, falseString))
				{
					return false;
				}
				if (falseString == null && trueString == null)
				{
					if (bool.TryParse(strVal, out bool b))
					{
						return b;
					}
					if (int.TryParse(strVal, NumberStyles.None, culture, out int v))
					{
						return v != 0;
					}
				}

				if (falseString == null && trueString != null) return false;
				if (trueString == null && falseString != null) return true;

				throw new InvalidCastException();
			case ExcelDataType.Error:
				var code = fi.ErrorCode;
				throw new ExcelFormulaException(ordinal, RowNumber, code);
		}
		throw new InvalidCastException();
	}

	/// <inheritdoc/>
	public override short GetInt16(int ordinal)
	{
		ValidateAccess();
		var i = GetInt32(ordinal);
		var s = (short)i;
		return s == i
			? s
			: throw new InvalidCastException();
	}

	/// <inheritdoc/>
	public override int GetInt32(int ordinal)
	{
		ValidateAccess();
		var type = GetExcelDataType(ordinal);
		switch (type)
		{
			case ExcelDataType.String:
				return int.Parse(GetString(ordinal), culture);
			case ExcelDataType.Numeric:
				var val = GetDouble(ordinal);
				var iVal = (int)val;
				if (iVal == val)
					return iVal;
				break;
		}

		throw new InvalidCastException();
	}

	/// <inheritdoc/>
	public override long GetInt64(int ordinal)
	{
		ValidateAccess();
		var type = GetExcelDataType(ordinal);
		switch (type)
		{
			case ExcelDataType.String:
				return long.Parse(GetString(ordinal));
			case ExcelDataType.Numeric:
				var val = GetDouble(ordinal);
				var iVal = (long)val;
				if (iVal == val)
					return iVal;
				break;
		}

		throw new InvalidCastException();
	}

	/// <inheritdoc/>
	public override decimal GetDecimal(int ordinal)
	{
		ValidateAccess();
		try
		{
			return (decimal)GetDouble(ordinal);
		}
		catch (OverflowException e)
		{
			throw new InvalidCastException(null, e);
		}
	}

	/// <inheritdoc/>
	public sealed override Guid GetGuid(int ordinal)
	{
		ValidateAccess();
		var val = this.GetString(ordinal);
		return Guid.TryParse(val, out var g)
			? g
			: throw new InvalidCastException();
	}

	/// <inheritdoc/>
	public sealed override byte GetByte(int ordinal)
	{
		ValidateAccess();
		var value = this.GetInt32(ordinal);
		var b = (byte)value;
		if (b == value)
		{
			return b;
		}
		else
		{
			throw new InvalidCastException();
		}
	}

	/// <inheritdoc/>
	public sealed override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length)
	{
		ValidateAccess();
		throw new NotSupportedException();
	}

	/// <inheritdoc/>
	public sealed override char GetChar(int ordinal)
	{
		ValidateAccess();
		var str = GetString(ordinal);
		if (str.Length == 1)
		{
			return str[0];
		}
		throw new InvalidCastException();
	}

	/// <inheritdoc/>
	public sealed override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length)
	{
		ValidateAccess();
		throw new NotSupportedException();
	}

	/// <inheritdoc/>
	public sealed override Stream GetStream(int ordinal)
	{
		ValidateAccess();
		throw new NotSupportedException();
	}

	/// <inheritdoc/>
	public sealed override TextReader GetTextReader(int ordinal)
	{
		ValidateAccess();
		throw new NotSupportedException();
	}

	/// <inheritdoc/>
	public override T GetFieldValue<T>(int ordinal)
	{
		ValidateAccess();
		var acc = Accessor<T>.Instance;
		return acc.GetValue(this, ordinal);
	}

	private protected enum State
	{
		None = 0,
		Initializing,
		// this state indicates that the next row is already in the field buffer
		// and should be returned as the next Read operation.
		Initialized,
		Open,
		End,
		Closed,
	}

	internal static double GetRKVal(int rk)
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
}
