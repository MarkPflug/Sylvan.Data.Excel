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
}
