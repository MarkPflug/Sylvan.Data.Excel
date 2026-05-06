#if !SPAN
using ReadonlyCharSpan = System.String;
#else
using ReadonlyCharSpan = System.ReadOnlySpan<char>;
using CharSpan = System.Span<char>;
#endif

namespace Sylvan.Data.Excel;

static class ExcelError
{
	static bool Equal(ReadonlyCharSpan l, ReadonlyCharSpan r)
	{
		if (l.Length != r.Length) return false;
		for (int i = 0; i < l.Length; i++)
		{
			if (l[i] != r[i])
				return false;
		}
		return true;
	}

	public static ExcelErrorCode GetErrorCode(ReadonlyCharSpan str)
	{
		if (Equal(str, "#DIV/0!"))
			return ExcelErrorCode.DivideByZero;
		if (Equal(str, "#VALUE!"))
			return ExcelErrorCode.Value;
		if (Equal(str, "#REF!"))
			return ExcelErrorCode.Reference;
		if (Equal(str, "#NAME?"))
			return ExcelErrorCode.Name;
		if (Equal(str, "#N/A"))
			return ExcelErrorCode.NotAvailable;
		if (Equal(str, "#NULL!"))
			return ExcelErrorCode.Null;
		if (Equal(str, "#NUM!"))
			return ExcelErrorCode.Number;

		throw new FormatException();
	}

	public static string GetErrorString(ExcelErrorCode code)
	{
		switch (code)
		{
			case ExcelErrorCode.DivideByZero:
				return "#DIV/0!";
			case ExcelErrorCode.Value:
				return "#VALUE!";
			case ExcelErrorCode.Reference:
				return "#REF!";
			case ExcelErrorCode.Name:
				return "#NAME?";
			case ExcelErrorCode.NotAvailable:
				return "#N/A";
			case ExcelErrorCode.Null:
				return "#NULL!";
			case ExcelErrorCode.Number:
				return "#NUM!";
			default:
				
				throw new InvalidDataException();
		}
	}
}

/// <summary>
/// Indicates the kind of error a formula produced.
/// </summary>
public enum ExcelErrorCode
{
	/// <summary>
	/// A null reference error.
	/// </summary>
	Null = 0,
	/// <summary>
	/// A division by zero error.
	/// </summary>
	DivideByZero = 7,
	/// <summary>
	/// A value error indicating a formula requires a numeric but was given a string.
	/// </summary>
	Value = 15,
	/// <summary>
	/// A reference error indicating a formula references a location that doesn't exist.
	/// </summary>
	Reference = 23,
	/// <summary>
	/// A name error indicating the formula references an unknown operation.
	/// </summary>
	Name = 29,
	/// <summary>
	/// A number error indicating the formula expected a number in a certain range.
	/// </summary>
	Number = 36,
	/// <summary>
	/// An error indicating the formula attempted to lookup a value that isn't available.
	/// </summary>
	NotAvailable = 42,
}

/// <summary>
/// An exception that is thrown when attempting to access a value in a cell that contains a function error.
/// </summary>
public sealed class ExcelFormulaException : Exception
{
	internal ExcelFormulaException(int col, int row, ExcelErrorCode code)
	{
		this.Row = row;
		this.Column = col;
		this.ErrorCode = code;
	}

	/// <summary>
	/// The row containing the error.
	/// </summary>
	public int Row { get; }

	/// <summary>
	/// The column containing the error.
	/// </summary>
	public int Column { get; }

	/// <summary>
	/// The error code indicating the kind of error.
	/// </summary>
	public ExcelErrorCode ErrorCode { get; }
}