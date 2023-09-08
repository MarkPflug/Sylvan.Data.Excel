namespace Sylvan.Data.Excel;

using System.IO.Compression;

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
	/// Creates a new ExcelDataWriterOptions instance when the default settings.
	/// </summary>
	public ExcelDataWriterOptions()
	{
		this.TruncateStrings = false;
		this.CompressionLevel = CompressionLevel.Fastest;
		this.OwnsStream = false;
	}

	/// <summary>
	/// Indicates if the ExcelDataWriter owns the output stream and handle disposal.
	/// </summary>
	public bool OwnsStream { get; set; }

	/// <summary>
	/// Indicates if string values should be truncated to the limit of Excel, which allows a maximum of 32k characters.
	/// </summary>
	/// <remarks>When false, an exception will be thrown if a string values exceeds the limit.</remarks>
	public bool TruncateStrings { get; set; }

	/// <summary>
	/// The compression level to use.
	/// </summary>
	public CompressionLevel CompressionLevel { get; set; }
}
