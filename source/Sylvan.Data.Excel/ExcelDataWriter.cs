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
		Dictionary<SharedStringEntry, string> dict;
		List<SharedStringEntry> entries;

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

			public override bool Equals(object obj)
			{
				return
					(obj is SharedStringEntry e)
					? this.Equals(e)
					: false;
			}

			public bool Equals(SharedStringEntry other)
			{
				return this.str.Equals(other.str);
			}
		}

		public string GetString(string str)
		{
			var entry = new SharedStringEntry(str);
			string idxStr;
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

	private protected SharedStringTable sharedStrings;

	/// <summary>
	/// Creates a new ExcelDataWriter.
	/// </summary>
	public static ExcelDataWriter Create(string file)
	{
		var stream = File.Create(file);
		var w = new XlsxDataWriter(stream);
		w.ownsStream = true;
		return w;
	}

	/// <summary>
	/// Creates a new ExcelDataWriter.
	/// </summary>
	public static ExcelDataWriter Create(Stream stream)
	{
		return new XlsxDataWriter(stream);
	}

	/// <inheritdoc/>
	public virtual void Dispose()
	{
		if (ownsStream)
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
