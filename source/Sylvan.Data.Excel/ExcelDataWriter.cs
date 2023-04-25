using Sylvan.Data.Excel.Xlsx;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Sylvan.Data.Excel;

/// <summary>
/// Writes data to excel files.
/// </summary>
public abstract class ExcelDataWriter : IDisposable
{
	bool ownsStream;
	readonly Stream stream; 
	private protected readonly bool truncateStrings;


	/// <summary>
	/// Creates a new ExcelDataWriter.
	/// </summary>
	public static ExcelDataWriter Create(string file, ExcelDataWriterOptions? options = null)
	{
		options = options ?? ExcelDataWriterOptions.Default;
		var type = ExcelDataReader.GetWorkbookType(file);
		var stream = File.Create(file);
		var w = Create(stream, type, options);
		w.ownsStream = true;
		return w;
	}

	/// <summary>
	/// Creates a new ExcelDataWriter.
	/// </summary>
	public static ExcelDataWriter Create(Stream stream, ExcelWorkbookType type, ExcelDataWriterOptions? options = null)
	{
		options = options ?? ExcelDataWriterOptions.Default;
		switch (type)
		{
			case ExcelWorkbookType.ExcelXml:
				{
					var w = new XlsxDataWriter(stream, options);
					return w;
				}
#if NET6_0_OR_GREATER
			case ExcelWorkbookType.ExcelBinary:
				{
					var w = new Xlsb.XlsbDataWriter(stream, options);
					return w;
				}
#endif
		}
		throw new NotSupportedException();
	}

	/// <inheritdoc/>
	public virtual void Dispose()
	{
		if (ownsStream)
			this.stream.Dispose();
	}

	private protected ExcelDataWriter(Stream stream, ExcelDataWriterOptions options)
	{
		this.stream = stream;
		this.truncateStrings = options.TruncateStrings;
	}

	/// <summary>
	/// Writes data to a new worksheet with the given name.
	/// </summary>
	/// <returns>The number of rows written.</returns>
	public abstract WriteResult Write(DbDataReader data, string? worksheetName = null);

	/// <summary>
	/// Writes data to a new worksheet with the given name.
	/// </summary>
	/// <returns>The number of rows written.</returns>
	public abstract Task<WriteResult> WriteAsync(DbDataReader data, string? worksheetName = null, CancellationToken cancel = default);

	/// <summary>
	/// A value indicating the result of the write operation.
	/// </summary>
	public readonly struct WriteResult
	{
		readonly int value;

		internal WriteResult(int val, bool complete)
		{
			this.value = complete ? val : -val;
		}

		/// <summary>
		/// Gets the number of rows written.
		/// </summary>
		public int RowsWritten => value < 0 ? -value : value;

		/// <summary>
		/// Indicates if all rows from the 
		/// </summary>
		public bool IsComplete => value >= 0;
	}
}
