using Sylvan.Data.Excel.Xlsx;
using System;
using System.Data.Common;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Sylvan.Data.Excel;

/// <summary>
/// Writes data to excel files.
/// </summary>
public abstract class ExcelDataWriter : 
	IDisposable
#if ASYNC
	, IAsyncDisposable
#endif
{

	bool isAsync;
#if ASYNC
	Stream? outputStream;
#endif

	bool ownsStream;
	readonly Stream stream;
	private protected readonly bool truncateStrings;


#if ASYNC

	/// <summary>
	/// Creates a new ExcelDataWriter.
	/// </summary>
	public static async Task<ExcelDataWriter> CreateAsync(string file, ExcelDataWriterOptions? options = null, CancellationToken cancel = default)
	{
		options = options ?? ExcelDataWriterOptions.Default;
		var type = ExcelDataReader.GetWorkbookType(file);
		var stream = File.Create(file);
		var w = await CreateAsync(stream, type, options, cancel).ConfigureAwait(false);
		w.ownsStream = true;
		return w;
	}

	/// <summary>
	/// Creates a new ExcelDataWriter.
	/// </summary>
	public static async Task<ExcelDataWriter> CreateAsync(Stream stream, ExcelWorkbookType type, ExcelDataWriterOptions? options = null, CancellationToken cancel = default)
	{
		options = options ?? ExcelDataWriterOptions.Default;
		ExcelDataWriter writer;
		var ms = new Sylvan.IO.PooledMemoryStream();
		switch (type)
		{
			case ExcelWorkbookType.ExcelXml:
				writer = new XlsxDataWriter(ms, options);
				break;
#if NET6_0_OR_GREATER
			case ExcelWorkbookType.ExcelBinary:
				writer = new Xlsb.XlsbDataWriter(ms, options);
				break;
#endif
			default:
				throw new NotSupportedException();
		}
		writer.isAsync = true;
		writer.outputStream = stream;
		// HACK: I want this method to be async to have symmetry with the `string filename` overload.
		await Task.CompletedTask.ConfigureAwait(false);
		return writer;
	}

#endif

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
		try
		{
			switch (type)
			{
				case ExcelWorkbookType.ExcelXml:
					return new XlsxDataWriter(stream, options);
#if NET6_0_OR_GREATER
				case ExcelWorkbookType.ExcelBinary:
					return new Xlsb.XlsbDataWriter(stream, options);
#endif
			}
			throw new NotSupportedException();
		} 
		catch
		{
			if (options?.OwnsStream == true)
			{
				stream.Dispose();
			}
			throw;
		}
	}

	/// <inheritdoc/>
	public virtual void Dispose()
	{
		if (isAsync)
		{
			throw new InvalidOperationException();
		}
		else
		{
			if (ownsStream)
			{
				this.stream.Dispose();
			}
		}
	}

#if ASYNC

	/// <inheritdoc/>
	public virtual async ValueTask DisposeAsync()
	{
		if (isAsync)
		{
			// outputStream should never be null here
			if (outputStream != null)
			{
				this.stream.Seek(0, SeekOrigin.Begin);
				await this.stream.CopyToAsync(this.outputStream!, CancellationToken.None).ConfigureAwait(true);
				if (ownsStream)
				{
					await this.outputStream!.DisposeAsync().ConfigureAwait(true);
				}
			}
		} 
		else
		{
			if (ownsStream)
			{
				await this.stream.DisposeAsync().ConfigureAwait(true);
			}
		}
	}
#endif

	private protected ExcelDataWriter(Stream stream, ExcelDataWriterOptions options)
	{
		this.isAsync = false;
		this.ownsStream = options.OwnsStream;
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
