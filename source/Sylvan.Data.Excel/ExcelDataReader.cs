using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Data.Common;
using System.IO;

namespace Sylvan.Data.Excel;
/// <summary>
/// A DbDataReader implementation that reads data from an Excel file.
/// </summary>
public abstract class ExcelDataReader : DbDataReader, IDisposable, IDbColumnSchemaGenerator
{
	static ReadOnlyCollection<DbColumn> EmptySchema = new ReadOnlyCollection<DbColumn>(Array.Empty<DbColumn>());

	private protected IExcelSchemaProvider schema;
	int fieldCount;
	private protected ReadOnlyCollection<DbColumn> columnSchema = EmptySchema;

	/// <inheritdoc/>
	public sealed override Type GetFieldType(int ordinal)
	{
		AssertRange(ordinal);
		if (ordinal < fieldCount)
		{
			return this.columnSchema[ordinal].DataType;
		}
		return typeof(object);
	}

	/// <inheritdoc/>
	public sealed override DataTable GetSchemaTable()
	{
		return SchemaTable.GetSchemaTable(this.GetColumnSchema());
	}

	private protected ExcelDataReader(IExcelSchemaProvider schema)
	{
		this.schema = schema;
	}

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

		var s = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read, 1);

		return Create(s, type, options);
	}

	/// <summary>
	/// Gets the number of fields in the current row.
	/// This may be different than FieldCount.
	/// </summary>
	public abstract int RowFieldCount { get; }

	/// <summary>
	/// Gets the maximum number of fields supported by the
	/// file.
	/// </summary>
	public abstract int MaxFieldCount { get; }

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
				return XlsWorkbookReader.CreateAsync(stream, options).GetAwaiter().GetResult();
			case ExcelWorkbookType.ExcelXml:
				return new XlsxWorkbookReader(stream, options);
			case ExcelWorkbookType.ExcelBinary:
				return new XlsbWorkbookReader(stream, options);
			default:
				throw new ArgumentException(nameof(fileType));
		}
	}

	static readonly Dictionary<string, ExcelWorkbookType> FileTypeMap = new(StringComparer.OrdinalIgnoreCase)
	{
		{ ".xls", ExcelWorkbookType.Excel },
		{ ".xlsx", ExcelWorkbookType.ExcelXml },
		{ ".xlsm", ExcelWorkbookType.ExcelXml },
		{ ".xlsb", ExcelWorkbookType.ExcelBinary },
	};

	/// <summary>
	/// Gets the type of an Excel workbook from the file name.
	/// </summary>
	public static ExcelWorkbookType GetWorkbookType(string filename)
	{
		var ext = Path.GetExtension(filename);
		return
			FileTypeMap.TryGetValue(ext, out var type)
			? type
			: 0;
	}

	/// <summary>
	/// Gets the number of worksheets in the workbook.
	/// </summary>
	public abstract int WorksheetCount { get; }

	/// <summary>
	/// Gets the name of the current worksheet.
	/// </summary>
	public abstract string? WorksheetName { get; }

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
	public abstract int RowCount { get; }

	/// <inheritdoc/>
	public sealed override int FieldCount => this.fieldCount;

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
	public abstract ExcelDataType GetExcelDataType(int ordinal);

	/// <summary>
	/// Gets the value as represented in excel.
	/// </summary>
	/// <remarks>
	/// Formula errors are returned as ExcelErrorCode values, rather than throwing an exception.
	/// </remarks>
	/// <param name="ordinal">The column ordinal to retrieve.</param>
	/// <returns>The value.</returns>
	public object GetExcelValue(int ordinal)
	{
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
	/// Gets the column schema
	/// </summary>
	public ReadOnlyCollection<DbColumn> GetColumnSchema()
	{
		return this.columnSchema;
	}

	/// <summary>
	/// Initializes the schema starting with the current row.
	/// </summary>
	public void InitializeSchema(IEnumerable<DbColumn> schema, bool useHeaders)
	{
		int i = 0;
		var cols = new List<DbColumn>();
		foreach (var col in schema)
		{
			var name = useHeaders ? this.GetString(i) : col.ColumnName;
			cols.Add(new ExcelColumn(name, i, col));
			i++;
		}

		this.columnSchema = new ReadOnlyCollection<DbColumn>(cols);
		this.fieldCount = columnSchema.Count;
	}

	private protected void LoadSchema(bool ordinalOnly)
	{
		var cols = new List<DbColumn>();
		var sheet = this.WorksheetName;
		if (sheet == null)
			throw new InvalidOperationException();
		for (int i = 0; i < RowFieldCount; i++)
		{
			string? header = ordinalOnly ? null : GetString(i);
			var col = schema.GetColumn(sheet, header, i);
			var ecs = new ExcelColumn(header, i, col);
			cols.Add(ecs);
		}
		this.columnSchema = new ReadOnlyCollection<DbColumn>(cols);
		this.fieldCount = columnSchema.Count;
	}

	/// <inheritdoc/>
	public sealed override int GetValues(object[] values)
	{
		var c = Math.Min(values.Length, this.FieldCount);

		for (int i = 0; i < c; i++)
		{
			values[i] = this.GetValue(i);
		}
		return c;
	}

	internal void AssertRange(int ordinal)
	{
		if((uint) ordinal >= MaxFieldCount)
		{
			throw new ArgumentOutOfRangeException(nameof(ordinal));
		}
	}

	/// <inheritdoc/>
	public sealed override object GetValue(int ordinal)
	{
		AssertRange(ordinal);
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
					return GetExcelValue(ordinal);
				}
				break;
		}
		throw new NotSupportedException();
	}

	/// <inheritdoc/>
	public sealed override IEnumerator GetEnumerator()
	{
		throw new NotSupportedException();
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

	internal abstract int DateEpochYear { get; }

	/// <summary>
	/// Gets the <see cref="ExcelErrorCode"/> of the error in the given cell.
	/// </summary>
	public abstract ExcelErrorCode GetFormulaError(int ordinal);

	internal ExcelFormulaException GetError(int ordinal)
	{
		return new ExcelFormulaException(ordinal, RowNumber, GetFormulaError(ordinal));
	}

	/// <summary>
	/// Gets the <see cref="ExcelFormat"/> of the format for the given cell.
	/// </summary>
	/// <param name="ordinal"></param>
	/// <returns></returns>
	public abstract ExcelFormat? GetFormat(int ordinal);

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
		var type = this.GetExcelDataType(ordinal);
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
				return TryGetDate(fmt, val, DateEpochYear, out var dt)
					? dt
					: throw new FormatException();
			case ExcelDataType.DateTime:
				return GetDateTimeValue(ordinal);
			case ExcelDataType.String:
			default:
				var str = GetString(ordinal);
				return DateTime.Parse(str);
		}
	}

	internal abstract DateTime GetDateTimeValue(int ordinal);

	static internal bool TryGetDate(ExcelFormat fmt, double value, int epoch, out DateTime dt)
	{
		dt = DateTime.MinValue;
		if (value < 61d && epoch == 1900)
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
				return false;
			}
		}
		else
		{
			value -= 1;
		}

		dt = new DateTime(epoch, 1, 1, 0, 0, 0, DateTimeKind.Unspecified).AddDays(value - 1d);
		return true;
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
	public abstract override string GetString(int ordinal);

	/// <inheritdoc/>
	public override float GetFloat(int ordinal)
	{
		return (float)GetDouble(ordinal);
	}

	/// <inheritdoc/>
	public override short GetInt16(int ordinal)
	{
		var i = GetInt32(ordinal);
		var s = (short)i;
		return s == i
			? s
			: throw new InvalidCastException();
	}

	/// <inheritdoc/>
	public override int GetInt32(int ordinal)
	{
		var type = GetExcelDataType(ordinal);
		switch (type)
		{
			case ExcelDataType.String:
				return int.Parse(GetString(ordinal));
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
		var val = this.GetString(ordinal);
		return Guid.TryParse(val, out var g)
			? g
			: throw new InvalidCastException();
	}

	/// <inheritdoc/>
	public sealed override byte GetByte(int ordinal)
	{
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
		throw new NotSupportedException();
	}

	/// <inheritdoc/>
	public sealed override char GetChar(int ordinal)
	{
		throw new NotSupportedException();
	}

	/// <inheritdoc/>
	public sealed override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length)
	{
		throw new NotSupportedException();
	}

	/// <inheritdoc/>
	public sealed override Stream GetStream(int ordinal)
	{
		throw new NotSupportedException();
	}

	/// <inheritdoc/>
	public sealed override TextReader GetTextReader(int ordinal)
	{
		throw new NotSupportedException();
	}

	private protected enum State
	{
		None = 0,
		Initializing,
		Initialized,
		Open,
		End,
		Closed,
	}
}
