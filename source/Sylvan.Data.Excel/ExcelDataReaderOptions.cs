using System.Globalization;

namespace Sylvan.Data.Excel;

/// <summary>
/// Options for controlling the behavior of an <see cref="ExcelDataReader"/>.
/// </summary>
public sealed class ExcelDataReaderOptions
{
	internal static readonly ExcelDataReaderOptions Default = new ExcelDataReaderOptions();

	/// <summary>
	/// Creates a new ExcelDataReaderOptions with the default values.
	/// </summary>
	public ExcelDataReaderOptions()
	{
		this.Schema = ExcelSchema.Default;
		this.Culture = CultureInfo.InvariantCulture;
		this.IgnoreEmptyTrailingRows = true;
	}

	/// <summary>
	/// Indicates that any trailing rows with empty cells should be ignored.
	/// Defaults to true.
	/// </summary>
	public bool IgnoreEmptyTrailingRows { get; set; }

	/// <summary>
	/// Gets or sets the schema for the data in the workbook.
	/// </summary>
	public IExcelSchemaProvider Schema { get; set; }

	/// <summary>
	/// Indicates if a cell will appear null or throw an ExcelFormulaException when accesing a cell containing a formula error.
	/// Defaults to false, which causes errors to be thrown.
	/// </summary>
	public bool GetErrorAsNull { get; set; }

	/// <summary>
	/// Indicates if hidden worksheets should be read, or skipped.
	/// Defaults to false, which skips hidden sheets.
	/// </summary>
	public bool ReadHiddenWorksheets { get; set; }

	/// <summary>
	/// The string which represents true values when reading boolean. Defaults to null.
	/// </summary>
	public string? TrueString { get; set; }

	/// <summary>
	/// The string which represents false values when reading boolean. Defaults to null.
	/// </summary>
	public string? FalseString { get; set; }

	/// <summary>
	/// A format string used to parse DateTime values.
	/// </summary>
	/// <remarks>
	/// This is only used in the very rare case that a date value is stored as a string
	/// in Excel, and is being accessed with GetDateTime() accessor.
	/// </remarks>
	public string? DateTimeFormat { get; set; }

	/// <summary>
	/// The culture to use when parsing values. 
	/// This is only used when accessing and converting values stored as a string.
	/// </summary>
	public CultureInfo Culture { get; set; }

	/// <summary>
	/// Indicates that the data stream should be disposed when the reader is disposed.
	/// </summary>
	public bool OwnsStream { get; set; }
}
