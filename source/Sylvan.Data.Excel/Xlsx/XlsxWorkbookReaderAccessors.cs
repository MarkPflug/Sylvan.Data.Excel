using System;
using System.Globalization;

#if SPAN
using ReadonlyCharSpan = System.ReadOnlySpan<char>;
using CharSpan = System.Span<char>;
#else
using ReadonlyCharSpan = System.String;
#endif

namespace Sylvan.Data.Excel;

partial class XlsxWorkbookReader
{
	private protected override bool GetBooleanValue(in FieldInfo fi, int ordinal)
	{
		if (fi.valueLen == 1)
		{
			return valuesBuffer[ordinal * ValueBufferElementSize] != '0';
		}
		throw new FormatException(); //?
	}

	private protected override double GetDoubleValue(in FieldInfo fi, int ordinal)
	{
#if SPAN
		var span = valuesBuffer.AsSpan(ordinal * ValueBufferElementSize, fi.valueLen);
		var value = double.Parse(span, NumberStyles.Float, CultureInfo.InvariantCulture);
#else
		var str = new string(valuesBuffer, ordinal * ValueBufferElementSize, fi.valueLen);
		var value = double.Parse(str, CultureInfo.InvariantCulture);
#endif
		return value;
	}

	public override bool GetBoolean(int ordinal)
	{
		ValidateAccess();
		ref readonly var fi = ref this.GetFieldValue(ordinal);
		switch (fi.type)
		{
			case FieldType.Boolean:
				return GetBooleanValue(in fi, ordinal);
			case FieldType.Numeric:
				return this.GetDoubleValue(in fi, ordinal) != 0;
			case FieldType.String:
			case FieldType.SharedString:

				var col = (uint)ordinal < this.columnSchema.Length ? this.columnSchema[ordinal] : null;

				var trueString = col?.TrueString ?? this.trueString;
				var falseString = col?.FalseString ?? this.falseString;

				var strVal = GetStringValue(in fi, ordinal);
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
					if (int.TryParse(strVal, NumberStyles.None, CultureInfo.InvariantCulture, out int v))
					{
						return v != 0;
					}
				}

				if (falseString == null && trueString != null) return false;
				if (trueString == null && falseString != null) return true;

				throw new InvalidCastException();
			case FieldType.Error:
				var code = fi.ErrorCode;
				throw new ExcelFormulaException(ordinal, RowNumber, code);
		}
		throw new InvalidCastException();
	}

	int GetValueInt(int ordinal)
	{
		ref readonly var fi = ref GetFieldValue(ordinal);
		if (TryParse(valuesBuffer.AsSpan().ToParsable(ordinal * ValueBufferElementSize, fi.valueLen), out int value))
		{
			return value;
		}
		throw new FormatException();
	}

	private protected override int GetSharedStringIndex(in FieldInfo fi, int ordinal)
	{
		return GetValueInt(ordinal);
	}

	public override double GetDouble(int ordinal)
	{
		ValidateAccess();
		ref readonly var cell = ref GetFieldValue(ordinal);
		switch (cell.type)
		{
			case FieldType.String:
			case FieldType.SharedString:
				return double.Parse(GetStringValue(in cell, ordinal), CultureInfo.InvariantCulture);
			case FieldType.Numeric:
				return GetDoubleValue(in cell, ordinal);
			case FieldType.Error:
				throw Error(ordinal);
		}

		throw new InvalidCastException();
	}

	private protected override ExcelFormulaException Error(int ordinal)
	{
		ref readonly var cell = ref GetFieldValue(ordinal);
		return new ExcelFormulaException(ordinal, RowNumber, GetFormulaError(ordinal));
	}

	public override ExcelErrorCode GetFormulaError(int ordinal)
	{
		ValidateAccess();
		var cell = GetFieldValue(ordinal);
		if (cell.type == FieldType.Error)
		{
#if SPAN
			var str = valuesBuffer.AsSpan(ordinal * ValueBufferElementSize, cell.valueLen);
#else
			var str = new String(valuesBuffer, ordinal * ValueBufferElementSize, cell.valueLen);
#endif
			return GetErrorCode(str);
		}			
		throw new InvalidOperationException();
	}
}
