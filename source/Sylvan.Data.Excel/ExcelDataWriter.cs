using Sylvan.Data.Excel.Xlsx;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;

namespace Sylvan.Data.Excel;

/// <summary>
/// Writes data to excel files.
/// </summary>
public abstract class ExcelDataWriter : IDisposable
{
	private protected class SharedStringTable
	{
		Dictionary<SharedStringEntry, int> dict;
		List<SharedStringEntry> entries;

		public int UniqueCount => entries.Count;

		public string this[int idx] => entries[idx].str;

		public SharedStringTable()
		{
			this.dict = new Dictionary<SharedStringEntry, int>();
			this.entries = new List<SharedStringEntry>();
		}

		struct SharedStringEntry
		{
			public string str;
			public bool isFormatted;

			public SharedStringEntry(string str, bool isFormatted = false)
			{
				this.str = str;
				this.isFormatted = isFormatted;
			}

			public override int GetHashCode()
			{
#if NETSTANDARD2_1_OR_GREATER
				return HashCode.Combine(str, isFormatted);
#else
				throw new NotImplementedException();
#endif
			}
		}

		public int GetString(string str)
		{
			var entry = new SharedStringEntry(str);
			int idx;
			if (!dict.TryGetValue(entry, out idx))
			{
				idx = this.entries.Count;
				this.entries.Add(entry);
				this.dict.Add(entry, idx);
			}
			return idx;
		}
	}

	readonly Stream stream;

	private protected SharedStringTable sharedStrings;

	/// <summary>
	/// Creates a new ExcelDataWriter.
	/// </summary>
	public static ExcelDataWriter Create(string file)
	{
		var stream = File.Create(file);
		return new XlsxDataWriter(stream);
	}

	/// <inheritdoc/>
	public virtual void Dispose()
	{
		this.stream.Dispose();
	}

	private protected ExcelDataWriter(Stream stream)
	{
		this.stream = stream;
		this.sharedStrings = new SharedStringTable();
	}

	/// <summary>
	/// Writes data to a new worksheet with the given name.
	/// </summary>
	public abstract void Write(string worksheetName, DbDataReader data);
}
