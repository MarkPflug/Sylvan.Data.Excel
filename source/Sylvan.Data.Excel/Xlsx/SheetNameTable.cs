using System;
using System.Xml;

namespace Sylvan.Data.Excel.Xlsx
{
	// these name tables avoid having to compute hashes on the
	// most common element names, and ensures that the strings
	// map the the const string so that equality will be ref
	// equality to the const values in our code.
	sealed class SheetNameTable : NameTable
	{
		public override string Add(char[] key, int start, int len)
		{
			return Get(key.AsSpan(start, len)) ?? base.Add(key, start, len);
		}

		public override string Add(string key)
		{
			return Get(key.AsSpan()) ?? base.Add(key);
		}

		public override string? Get(char[] key, int start, int len)
		{
			return Get(key.AsSpan(start, len));
		}

		public override string? Get(string value)
		{
			return Get(value.AsSpan());
		}

		public string? Get(ReadOnlySpan<char> value)
		{
			switch (value.Length)
			{
				case 0:
					return string.Empty;
				case 1:
					switch (value[0])
					{
						case 'c': return "c";
						case 'r': return "r";
						case 't': return "t";
						case 's': return "s";
						case 'v': return "v";
					}
					break;
				case 2:
					if (value.SequenceEqual("is")) return "is";
					break;
				case 3:
					if (value.SequenceEqual("row")) return "row";
					if (value.SequenceEqual("ref")) return "ref";
					break;
				case 5:
					if (value.SequenceEqual("spans")) return "spans";
					break;
				case 9:
					if (value.SequenceEqual("dyDescent")) return "dyDescent";
					if (value.SequenceEqual("dimension")) return "dimension";
					if (value.SequenceEqual("sheetData")) return "sheetData";
					break;
			}
			return null;
		}
	}

	sealed class SharedStringsNameTable : NameTable
	{
		public override string Add(char[] key, int start, int len)
		{
			return Get(key.AsSpan(start, len)) ?? base.Add(key, start, len);
		}

		public override string Add(string key)
		{
			return Get(key.AsSpan()) ?? base.Add(key);
		}

		public override string? Get(char[] key, int start, int len)
		{
			return Get(key.AsSpan(start, len));
		}

		public override string? Get(string value)
		{
			return Get(value.AsSpan());
		}

		public string? Get(ReadOnlySpan<char> value)
		{
			switch (value.Length)
			{
				case 0:
					return string.Empty;
				case 1:
					if (value.SequenceEqual("t")) return "t";
					break;
				case 2:
					if (value.SequenceEqual("si")) return "si";
					break;
				case 3:
					if (value.SequenceEqual("sst")) return "sst";
					break;
			}
			return null;
		}
	}
}
