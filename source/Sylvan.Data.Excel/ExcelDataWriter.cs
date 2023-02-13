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
#if ASYNC_DISPOSE
	, IAsyncDisposable
#endif
{
	private protected class SharedStringTable
	{
		readonly Dictionary<SharedStringEntry, string> dict;
		readonly List<SharedStringEntry> entries;

		public int UniqueCount => entries.Count;

		public string this[int idx] => entries[idx].str;

		public SharedStringTable()
		{
			const int InitialSize = 128;
			this.dict = new Dictionary<SharedStringEntry, string>(InitialSize);
			this.entries = new List<SharedStringEntry>(InitialSize);
		}

		struct SharedStringEntry : IEquatable<SharedStringEntry>
		{
			public string str;
			public string idxStr;

			public SharedStringEntry(string str)
			{
				this.str = str;
				this.idxStr = "";
			}

			public override int GetHashCode()
			{
				return str.GetHashCode();
			}

			public override bool Equals(object? obj)
			{
				return (obj is SharedStringEntry e) && this.Equals(e);
			}

			public bool Equals(SharedStringEntry other)
			{
				return this.str.Equals(other.str);
			}
		}

		public string GetString(string str)
		{
			var entry = new SharedStringEntry(str);
			string? idxStr;
			if (!dict.TryGetValue(entry, out idxStr))
			{
				idxStr = this.entries.Count.ToString();
				this.entries.Add(entry);
				this.dict.Add(entry, idxStr);
			}
			return idxStr;
		}
	}

	bool ownsStream;
	readonly Stream stream;
	private protected readonly bool truncateStrings;
	Stream? userStream;
	bool isAsync = false;

	private protected SharedStringTable sharedStrings;

	/// <summary>
	/// Creates a new ExcelDataWriter to be used with asynchronous writing.
	/// </summary>
	public static async Task<ExcelDataWriter> CreateAsync(string file, ExcelDataWriterOptions? options = null)
	{
		options = options ?? ExcelDataWriterOptions.Default;
		var type = ExcelDataReader.GetWorkbookType(file);
		var stream = File.Create(file);
		var writer = await CreateAsync(stream, type, options).ConfigureAwait(false);
		writer.ownsStream = true;
		return writer;
	}

	/// <summary>
	/// Creates a new ExcelDataWriter to be used with asynchronous writing.
	/// </summary>
	public static Task<ExcelDataWriter> CreateAsync(Stream stream, ExcelWorkbookType type, ExcelDataWriterOptions? options = null)
	{
		// if the stream that is being written to might block (FileStream)
		// then write to a MemoryStream instead, and asynchronously
		// flush to the user stream during async dispose.
		options = options ?? ExcelDataWriterOptions.Default;
		Stream userStream = stream;
		Stream asyncStream =
			stream is MemoryStream ms
			? ms
			: new MemoryStream();
		
		switch (type)
		{
			case ExcelWorkbookType.ExcelXml:				
				var w = new XlsxDataWriter(asyncStream, options);
				w.isAsync = true;
				w.userStream = userStream;

				w.ownsStream = false;
				return Task.FromResult((ExcelDataWriter)w);
		}
		throw new NotSupportedException();
	}

	/// <summary>
	/// Creates a new ExcelDataWriter.
	/// </summary>
	public static ExcelDataWriter Create(string file, ExcelDataWriterOptions? options = null)
	{
		options = options ?? ExcelDataWriterOptions.Default;
		var type = ExcelDataReader.GetWorkbookType(file);
		switch (type)
		{
			case ExcelWorkbookType.ExcelXml:
				var stream = File.Create(file);
				var w = new XlsxDataWriter(stream, options);
				w.ownsStream = true;
				return w;
		}
		throw new NotSupportedException();
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
				return new XlsxDataWriter(stream, options);
		}
		throw new NotSupportedException();
	}

	private protected ExcelDataWriter(Stream stream, ExcelDataWriterOptions options)
	{
		this.stream = stream;
		this.sharedStrings = new SharedStringTable();
		this.truncateStrings = options.TruncateStrings;
	}

	/// <inheritdoc/>
	public virtual void Dispose()
	{
		if (isAsync)
		{
			// this call pattern would be incorrect.
			// DisposeAsync should have been called
			// but we should do our best to clean up
			if (userStream != null)
			{
				stream.Seek(0, SeekOrigin.Begin);
				stream.CopyTo(this.userStream);
			}
			if (ownsStream)
			{
				if (this.userStream != null)
				{
					this.userStream.Dispose();
				}
			}
		}
		else
		{
			this.stream.Dispose();
		}
	}

#if ASYNC_DISPOSE
	/// <inheritdoc/>
	public virtual async ValueTask DisposeAsync()
	{
		if (isAsync)
		{
			if (userStream != null)
			{
				stream.Seek(0, SeekOrigin.Begin);
				await stream.CopyToAsync(this.userStream).ConfigureAwait(false);
			}
			if (ownsStream)
			{
				if (this.userStream != null)
				{
					await this.userStream.DisposeAsync().ConfigureAwait(false);
				}
			}
		}
		else
		{
			await this.stream.DisposeAsync().ConfigureAwait(false);
		}
	}
#endif



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
