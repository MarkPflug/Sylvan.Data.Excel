#if !SPAN

using System;

namespace Sylvan;

readonly struct CharSpan
{
	readonly char[] buffer;
	readonly int offset;
	readonly int length;

	public CharSpan(char[] buffer)
		: this(buffer, 0, buffer.Length)
	{

	}

	public CharSpan(char[] buffer, int offset, int length)
	{
		if (offset < 0 || offset >= buffer.Length)
			throw new ArgumentOutOfRangeException(nameof(offset));
		var end = offset + length;
		if (end < 0 || end > buffer.Length)
			throw new ArgumentOutOfRangeException(nameof(length));
		this.buffer = buffer;
		this.offset = offset;
		this.length = length;
	}

	public int Length => this.length;

	public char this[int idx]
	{
		get
		{
			return this.buffer[offset + idx];
		}
		set
		{
			if ((uint)idx >= this.length)
				throw new IndexOutOfRangeException();
			this.buffer[offset + idx] = value;
		}
	}

	public CharSpan Slice(int offset)
	{
		var o = this.offset + offset;

		var length = this.length - o;
		if (o >= buffer.Length)
			throw new ArgumentOutOfRangeException(nameof(offset));

		return new CharSpan(this.buffer, o, length);
	}

	public CharSpan Slice(int offset, int length)
	{
		if (offset + length > this.length)
			throw new ArgumentOutOfRangeException(nameof(length));
		return new CharSpan(this.buffer, this.offset + offset, length);
	}

	public override string ToString()
	{
		return new string(buffer, offset, length);
	}
}

static class RunTimeCompatability
{
	public static CharSpan AsSpan(this char[] value)
	{
		return new CharSpan(value);
	}

	public static CharSpan AsSpan(this char[] value, int start, int length)
	{
		return new CharSpan(value, start, length);
	}

	public static char[] ToArray(this string value)
	{
		return value.ToCharArray();
	}

	public static string ToParsable(this CharSpan span)
	{
		return span.ToString();
	}

	public static string ToParsable(this CharSpan span, int offset, int length)
	{
		return span.Slice(offset, length).ToString();
	}

	public static CharSpan ToCharSpan(this CharSpan span, int offset, int length)
	{
		return span.Slice(offset, length);
	}

	public static CharSpan AllocateCharSpan(int length)
	{
		return new CharSpan(new char[length]);
	}
}
#else

using System;

namespace Sylvan;

static class RunTimeCompatability
{
	public static ReadOnlySpan<char> ToParsable(this ReadOnlySpan<char> span, int offset, int length)
	{
		return span.Slice(offset, length);
	}

	public static ReadOnlySpan<char> ToParsable(this Span<char> span, int offset, int length)
	{
		return span.Slice(offset, length);
	}
}

#endif
