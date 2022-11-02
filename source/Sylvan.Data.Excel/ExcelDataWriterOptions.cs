namespace Sylvan.Data.Excel;

/// <summary>
/// Options for controlling the behavior of an <see cref="ExcelDataWriter"/>.
/// </summary>
public sealed class ExcelDataWriterOptions
{
	internal static readonly ExcelDataWriterOptions Default = new ExcelDataWriterOptions();

	// TODO: compression option?
	// formats? Date? boolean?
	// shared string behavior?
	// binary as hex/basxe64

	/// <summary>
	/// Indicates if string values should be truncated to the limit of Excel, which allows a maximum of 32k characters.
	/// </summary>
	public bool TruncateStrings { get; set; }

}
