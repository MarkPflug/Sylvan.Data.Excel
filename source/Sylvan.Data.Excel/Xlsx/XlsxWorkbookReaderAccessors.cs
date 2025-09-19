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

	bool GetBooleanRaw(ref readonly FieldInfo fi, int ordinal)
	{
		if (fi.valueLen == 1)
		{
			return valuesBuffer[ordinal * ValueBufferElementSize] != '0';
		}
		throw new FormatException(); //?
	}

	public override bool GetBoolean(int ordinal)
	{
		ValidateAccess();
		ref readonly var fi = ref this.GetFieldValue(ordinal);
		switch (fi.type)
		{
			case FieldType.Boolean:
				return GetBooleanRaw(in fi, ordinal);
			case FieldType.Numeric:
				return this.GetDoubleRaw(ordinal) != 0;
			case FieldType.String:
			case FieldType.SharedString:

				var col = (uint)ordinal < this.columnSchema.Length ? this.columnSchema[ordinal] : null;

				var trueString = col?.TrueString ?? this.trueString;
				var falseString = col?.FalseString ?? this.falseString;

				var strVal = GetStringRaw(ordinal, in fi);
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

	public override string GetString(int ordinal)
	{
		ValidateAccess();

		ref readonly FieldInfo fi = ref GetFieldValue(ordinal);
		if (ordinal >= MaxFieldCount)
		{
			throw new ArgumentOutOfRangeException(nameof(ordinal));
		}

		switch (fi.type)
		{
			case FieldType.Error:
				if (errorAsNull)
				{
					return string.Empty;
				}
				throw GetError(ordinal);
			case FieldType.Boolean:
				var boolVal = GetBooleanRaw(in fi, ordinal);
				return boolVal ? bool.TrueString : bool.FalseString;
			case FieldType.Numeric:
				var numValue = GetDoubleRaw(ordinal);
				return FormatVal(fi.xfIdx, numValue);
			case FieldType.String:
				return fi.strValue ?? string.Empty;
			case FieldType.SharedString:
				return GetStringRaw(ordinal, fi);
			case FieldType.Null:
				return string.Empty;
		}
		throw new NotSupportedException();
	}


	string GetStringRaw(int ordinal, in FieldInfo fi)
	{
		return (fi.type == FieldType.SharedString ? GetSharedStringRaw(in fi, ordinal) : fi.strValue) ?? string.Empty;

	}

	private protected override string GetSharedStringRaw(ref readonly FieldInfo fi, int ordinal)
	{
		var ssIdx = GetValueInt(ordinal);
		return GetSharedString(ssIdx) ?? "";
	}

	public override double GetDouble(int ordinal)
	{
		ValidateAccess();
		ref readonly var cell = ref GetFieldValue(ordinal);
		switch (cell.type)
		{
			case FieldType.String:
				return double.Parse(GetStringRaw(ordinal, in cell), culture);
			case FieldType.SharedString:
				return double.Parse(GetSharedStringRaw(in cell, ordinal), culture);
			case FieldType.Numeric:
				return GetDoubleRaw(ordinal);
			case FieldType.Error:
				throw Error(ordinal);
		}

		throw new InvalidCastException();
	}

	double GetDoubleRaw(int ordinal)
	{
		ref readonly var cell = ref GetFieldValue(ordinal);
#if SPAN
		var span = valuesBuffer.AsSpan(ordinal * ValueBufferElementSize, cell.valueLen);
		var value = double.Parse(span, NumberStyles.Float, CultureInfo.InvariantCulture);
#else
		var str = new string(valuesBuffer, ordinal * ValueBufferElementSize, cell.valueLen);
		var value = double.Parse(str, CultureInfo.InvariantCulture);
#endif
		return value;
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
