#if !NETSTANDARD2_1_OR_GREATER

using ReadonlyCharSpan = System.String;
using CharSpan = System.Text.StringBuilder;

namespace Sylvan;

static class RunTimeCompatability
{
	public static ReadonlyCharSpan AsSpan(this char[] value)
	{
		return new ReadonlyCharSpan(value);
	}

	public static ReadonlyCharSpan AsSpan(this char[] value, int start, int length)
	{
		return new ReadonlyCharSpan(value, start, length);
	}

	public static char[] ToArray(this string value)
	{
		return value.ToCharArray();
	}
	public static ReadonlyCharSpan Slice(this ReadonlyCharSpan value, int start, int length)
	{
		return value.Substring(start, length);
	}
	public static ReadonlyCharSpan Slice(this ReadonlyCharSpan value, int start)
	{
		return value.Substring(start);
	}

	public static ReadonlyCharSpan Slice(this CharSpan value, int start, int length)
	{
		return value.ToString(start, length);
	}

	public static ReadonlyCharSpan Slice(this CharSpan value, int start)
	{
		return value.ToString(start, value.Length - start);
	}

	public static CharSpan AllocateCharSpan(int length)
	{
		var charSpan = new CharSpan(length);
		charSpan.Append(new char[length]);
		return charSpan;
	}
}

#endif
