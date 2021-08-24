using System;
using System.Collections;
using System.Collections.ObjectModel;
using System.Data.Common;
using System.IO;

namespace Sylvan.Data.Excel
{
	public enum ExcelWorkbookType
	{
		Excel,
		OpenExcel,
		OpenExcelBinary,
	}

	public enum ExcelDataType
	{
		Null = 0,
		Numeric,
		String,
		Boolean,
		Error,
	}

	public sealed class ExcelDataReaderOptions
	{
		internal static readonly ExcelDataReaderOptions Default = new ExcelDataReaderOptions();

		public ExcelDataReaderOptions()
		{
			this.Schema = ExcelSchema.Default;
			this.GetNullAsEmptyString = true;
		}

		public IExcelSchemaProvider Schema { get; set; }

		/// <summary>
		/// Indicates if GetString will return an emtpy string
		/// when the underlying value is null. If false, GetString
		/// will throw an exception when the underlying value is null.
		/// The default is true.
		/// </summary>
		public bool GetNullAsEmptyString { get; set; }

		/// <summary>
		/// Indicates if GetString will throw an ExcelFormulaException or return a string value
		/// when accesing a cell containing a formula error.
		/// </summary>
		public bool GetErrorAsNull { get; set; }
	}

	public abstract class ExcelDataReader : DbDataReader, IDisposable, IDbColumnSchemaGenerator
	{
		public static long counter = 0;

		public static ExcelDataReader Create(string filename, ExcelDataReaderOptions? options = null)
		{
			options = options ?? ExcelDataReaderOptions.Default;

			var ext = Path.GetExtension(filename);

			if (StringComparer.OrdinalIgnoreCase.Equals(".xls", ext))
			{
				var s = File.OpenRead(filename);
				//var s = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read, 0x100000);
				var pkg = new Ole2Package(s);
				var part = pkg.GetEntry("Workbook\0");
				if (part == null)
					throw new InvalidDataException();
				var ps = part.Open();
				return XlsWorkbookReader.CreateAsync(ps, options).GetAwaiter().GetResult();
			}

			if (StringComparer.OrdinalIgnoreCase.Equals(".xlsx", ext))
			{
				var s = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read, 0x10000);
				return new XlsxWorkbookReader(s, options);
			}
			throw new NotSupportedException();
		}

		/// <summary>
		/// Gets the number of worksheets in the workbook.
		/// </summary>
		public abstract int WorksheetCount { get; }

		/// <summary>
		/// Gets the name of the current worksheet.
		/// </summary>
		public abstract string WorksheetName { get; }

		/// <summary>
		/// Gets the number of rows in the current sheet.
		/// </summary>
		/// <remarks>
		/// Can return -1 to indicate that the number of rows is unknown.
		/// </remarks>
		public abstract int RowCount { get; }

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

		public abstract ReadOnlyCollection<DbColumn> GetColumnSchema();

		public sealed override int GetValues(object[] values)
		{
			var c = Math.Min(values.Length, this.FieldCount);

			for (int i = 0; i < c; i++)
			{
				values[i] = this.GetValue(i);
			}
			return c;
		}

		public sealed override IEnumerator GetEnumerator()
		{
			throw new NotSupportedException();
		}

		public sealed override int Depth => 0;

		public sealed override object this[int ordinal] => this.GetValue(ordinal);
		public sealed override object this[string name] => this.GetValue(this.GetOrdinal(name));

		public sealed override string GetDataTypeName(int ordinal)
		{
			return this.GetFieldType(ordinal).Name;
		}

		public sealed override int RecordsAffected => 0;

		public sealed override bool HasRows => this.RowCount != 0;

		internal abstract int DateEpochYear { get; }

		public abstract ExcelErrorCode GetFormulaError(int ordinal);

		public abstract ExcelFormat? GetFormat(int ordinal);

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
					return TryGetDate(val, DateEpochYear, out var dt)
						? dt
						: throw new FormatException();
				case ExcelDataType.String:
				default:
					var str = GetString(ordinal);
					return DateTime.Parse(str);
			}
		}

		internal static bool TryGetDate(double value, int epoch, out DateTime dt)
		{
			dt = DateTime.MinValue;
			if (value < 61d && epoch == 1900)
			{
				if (value < 1)
				{
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

		public override float GetFloat(int ordinal)
		{
			return (float)GetDouble(ordinal);
		}

		public override short GetInt16(int ordinal)
		{
			var i = GetInt32(ordinal);
			var s = (short)i;
			return s == i
				? s
				: throw new InvalidCastException();
		}

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

		public sealed override Guid GetGuid(int ordinal)
		{
			var val = this.GetString(ordinal);
			return Guid.TryParse(val, out var g)
				? g
				: throw new InvalidCastException();
		}

		public sealed override byte GetByte(int ordinal)
		{
			throw new NotSupportedException();
		}

		public sealed override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length)
		{
			throw new NotSupportedException();
		}

		public sealed override char GetChar(int ordinal)
		{
			throw new NotSupportedException();
		}

		public sealed override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length)
		{
			throw new NotSupportedException();
		}

		public sealed override Stream GetStream(int ordinal)
		{
			throw new NotSupportedException();
		}

		public sealed override TextReader GetTextReader(int ordinal)
		{
			throw new NotSupportedException();
		}
	}
}
