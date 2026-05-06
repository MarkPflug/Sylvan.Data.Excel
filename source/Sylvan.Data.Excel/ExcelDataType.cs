namespace Sylvan.Data.Excel;

/// <summary>
/// Represents in internal data types supported by Excel.
/// </summary>
public enum ExcelDataType
{
	/// <summary>
	/// A cell that contains no value.
	/// </summary>
	Null = 0,
	/// <summary>
	/// A numeric value. These cells might have a style applied that formats as a date or time value.
	/// </summary>
	Numeric,
	/// <summary>
	/// A DateTime value. This is used in "strict" .xlsx files.
	/// </summary>
	DateTime,
	/// <summary>
	/// A text field.
	/// </summary>
	String,
	/// <summary>
	/// A formula cell that contains a boolean.
	/// </summary>
	Boolean,
	/// <summary>
	/// A formula cell that contains an error.
	/// </summary>
	Error,
}
