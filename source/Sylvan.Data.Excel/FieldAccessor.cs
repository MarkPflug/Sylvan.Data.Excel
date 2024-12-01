using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Sylvan.Data.Excel;

static class Accessor<T>
{
	public static IFieldAccessor<T> Instance = GetAccessor();

	static IFieldAccessor<T> GetAccessor()
	{
		var acc = ExcelDataAccessor.Instance as IFieldAccessor<T>;
		if (acc == null)
		{
			if (typeof(T).IsEnum)
			{
				return EnumAccessor<T>.Instance;
			}
			throw new NotSupportedException(); // TODO: exception type?
		}
		return acc;
	}
}

delegate bool TryParse<T>(string str, bool ignoreCase, out T value);

static class EnumParse
{
	internal static MethodInfo? GenericSpanParseMethod;

	static MethodInfo? GetGenericMethod()
	{
		return
			typeof(Enum)
			.GetMethods()
			.Where(m =>
			{
				if (m.Name != "TryParse")
					return false;
				var p = m.GetParameters();
				return p.Count() == 3 && p[0].ParameterType == typeof(string);
			}
			)
			.SingleOrDefault();
	}

	static EnumParse()
	{
		GenericSpanParseMethod = GetGenericMethod();
	}
}

sealed class EnumAccessor<T> : IFieldAccessor<T>
{
	internal static EnumAccessor<T> Instance = new EnumAccessor<T>();

	internal static TryParse<T>? Parser;

	static EnumAccessor()
	{
		Parser = null;
		var method = EnumParse.GenericSpanParseMethod;
		if (method != null)
		{
			var gm = method.MakeGenericMethod(new[] { typeof(T) });
			Parser = (TryParse<T>)gm.CreateDelegate(typeof(TryParse<T>));
		}
	}

	public T GetValue(ExcelDataReader reader, int ordinal)
	{
		var parser = Parser;
		if (parser == null)
		{
			throw new NotSupportedException();
		}
		var span = reader.GetString(ordinal);
		if (span.Length == 0) return default!;
		return
			parser(span, true, out T value)
			? value
			: throw new InvalidCastException();
	}
}

// these accessors support the GetFieldValue<T> generic accessor method.
// most of them defer to ExcelDataReader.GetXXX methods.

interface IFieldAccessor<T>
{
	T GetValue(ExcelDataReader reader, int ordinal);
}

interface IFieldAccessor
{
	object GetValueAsObject(ExcelDataReader reader, int ordinal);
}

interface IFieldRangeAccessor<T>
{
	long GetRange(ExcelDataReader reader, long dataOffset, int ordinal, T[] buffer, int bufferOffset, int length);
}

abstract class FieldAccessor<T> : IFieldAccessor<T>, IFieldAccessor
{
	public object GetValueAsObject(ExcelDataReader reader, int ordinal)
	{
		return (object)this.GetValue(reader, ordinal)!;
	}

	public abstract T GetValue(ExcelDataReader reader, int ordinal);
}

sealed class StringAccessor : FieldAccessor<string>
{
	internal static readonly StringAccessor Instance = new StringAccessor();

	public override string GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetString(ordinal);
	}
}

sealed class BooleanAccessor : FieldAccessor<bool>
{

	internal static readonly BooleanAccessor Instance = new BooleanAccessor();

	public override bool GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetBoolean(ordinal);
	}
}

sealed class CharAccessor : FieldAccessor<char>
{
	internal static readonly CharAccessor Instance = new CharAccessor();

	public override char GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetChar(ordinal);
	}
}

sealed class ByteAccessor : FieldAccessor<byte>
{
	internal static readonly ByteAccessor Instance = new ByteAccessor();

	public override byte GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetByte(ordinal);
	}
}

sealed class Int16Accessor : FieldAccessor<short>
{
	internal static readonly Int16Accessor Instance = new Int16Accessor();

	public override short GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetInt16(ordinal);
	}
}

sealed class Int32Accessor : FieldAccessor<int>
{
	internal static readonly Int32Accessor Instance = new Int32Accessor();

	public override int GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetInt32(ordinal);
	}
}

sealed class Int64Accessor : FieldAccessor<long>
{
	internal static readonly Int64Accessor Instance = new Int64Accessor();

	public override long GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetInt64(ordinal);
	}
}

sealed class SingleAccessor : FieldAccessor<float>
{
	internal static readonly SingleAccessor Instance = new SingleAccessor();

	public override float GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetFloat(ordinal);
	}
}

sealed class DoubleAccessor : FieldAccessor<double>
{
	internal static readonly DoubleAccessor Instance = new DoubleAccessor();

	public override double GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetDouble(ordinal);
	}
}

sealed class DecimalAccessor : FieldAccessor<decimal>
{
	internal static readonly DecimalAccessor Instance = new DecimalAccessor();

	public override decimal GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetDecimal(ordinal);
	}
}

sealed class DateTimeAccessor : FieldAccessor<DateTime>
{
	internal static readonly DateTimeAccessor Instance = new DateTimeAccessor();

	public override DateTime GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetDateTime(ordinal);
	}
}

sealed class TimeSpanAccessor : FieldAccessor<TimeSpan>
{
	internal static readonly TimeSpanAccessor Instance = new TimeSpanAccessor();

	public override TimeSpan GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetTimeSpan(ordinal);
	}
}

#if DATE_ONLY

sealed class DateOnlyAccessor : FieldAccessor<DateOnly>
{
	internal static readonly DateOnlyAccessor Instance = new DateOnlyAccessor();

	public override DateOnly GetValue(ExcelDataReader reader, int ordinal)
	{
		var dt = reader.GetDateTime(ordinal);
		return new DateOnly(dt.Year, dt.Month, dt.Day);
	}
}

sealed class TimeOnlyAccessor : FieldAccessor<TimeOnly>
{
	internal static readonly TimeOnlyAccessor Instance = new TimeOnlyAccessor();

	public override TimeOnly GetValue(ExcelDataReader reader, int ordinal)
	{
		var dt = reader.GetDateTime(ordinal);
		return TimeOnly.FromDateTime(dt);
	}
}

#endif


sealed class GuidAccessor : FieldAccessor<Guid>
{
	internal static readonly GuidAccessor Instance = new GuidAccessor();

	public override Guid GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetGuid(ordinal);
	}
}

//sealed class StreamAccessor : FieldAccessor<Stream>
//{
//	internal static readonly StreamAccessor Instance = new StreamAccessor();

//	public override Stream GetValue(ExcelDataReader reader, int ordinal)
//	{
//		return reader.GetStream(ordinal);
//	}
//}

//sealed class TextReaderAccessor : FieldAccessor<TextReader>
//{
//	internal static readonly TextReaderAccessor Instance = new TextReaderAccessor();

//	public override TextReader GetValue(ExcelDataReader reader, int ordinal)
//	{
//		return reader.GetTextReader(ordinal);
//	}
//}

//sealed class BytesAccessor : FieldAccessor<byte[]>
//{
//	internal static readonly BytesAccessor Instance = new BytesAccessor();

//	public override byte[] GetValue(ExcelDataReader reader, int ordinal)
//	{
//		var len = reader.GetBytes(ordinal, 0, null, 0, 0);
//		var buffer = new byte[len];
//		reader.GetBytes(ordinal, 0, buffer, 0, buffer.Length);
//		return buffer;
//	}
//}

//sealed class CharsAccessor : FieldAccessor<char[]>
//{
//	internal static readonly CharsAccessor Instance = new CharsAccessor();

//	public override char[] GetValue(ExcelDataReader reader, int ordinal)
//	{
//		var len = reader.GetChars(ordinal, 0, null, 0, 0);
//		var buffer = new char[len];
//		reader.GetChars(ordinal, 0, buffer, 0, buffer.Length);
//		return buffer;
//	}
//}

sealed partial class ExcelDataAccessor :
	IFieldAccessor<string>,
	IFieldAccessor<bool>,
	IFieldAccessor<char>,
	IFieldAccessor<byte>,
	IFieldAccessor<short>,
	IFieldAccessor<int>,
	IFieldAccessor<long>,
	IFieldAccessor<float>,
	IFieldAccessor<double>,
	IFieldAccessor<decimal>,
	IFieldAccessor<DateTime>,
	IFieldAccessor<TimeSpan>,
#if DATE_ONLY
	IFieldAccessor<DateOnly>,
	IFieldAccessor<TimeOnly>,
#endif
	IFieldAccessor<Guid>,
	IFieldAccessor<Stream>,
	IFieldAccessor<TextReader>,
	IFieldAccessor<byte[]>,
	IFieldAccessor<char[]>,
	IFieldRangeAccessor<byte>,
	IFieldRangeAccessor<char>
{
	internal static readonly ExcelDataAccessor Instance = new ExcelDataAccessor();

	internal static readonly Dictionary<Type, IFieldAccessor> Accessors;

	static ExcelDataAccessor()
	{
		Accessors = new Dictionary<Type, IFieldAccessor>
		{
			{typeof(string), StringAccessor.Instance },
			{typeof(bool), BooleanAccessor.Instance },
			{typeof(char), CharAccessor.Instance },
			{typeof(byte), ByteAccessor.Instance },
			{typeof(short), Int16Accessor.Instance },
			{typeof(int), Int32Accessor.Instance },
			{typeof(long), Int64Accessor.Instance },
			{typeof(float), SingleAccessor.Instance },
			{typeof(double), DoubleAccessor.Instance },
			{typeof(decimal), DecimalAccessor.Instance },
			{typeof(DateTime), DateTimeAccessor.Instance },
			{typeof(TimeSpan), TimeSpanAccessor.Instance },
#if DATE_ONLY
			{typeof(DateOnly), DateOnlyAccessor.Instance },
			{typeof(TimeOnly), TimeOnlyAccessor.Instance },
#endif
			{typeof(Guid), GuidAccessor.Instance },

			// TODO: add support for the following types?
			//{typeof(Stream), StreamAccessor.Instance },
			//{typeof(TextReader), TextReaderAccessor.Instance },
			//{typeof(byte[]), BytesAccessor.Instance },
			//{typeof(char[]), CharsAccessor.Instance },
		};
	}

	internal static IFieldAccessor? GetAccessor(Type type)
	{
		return Accessors.TryGetValue(type, out var acc) ? acc : null;
	}

	string IFieldAccessor<string>.GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetString(ordinal);
	}

	long IFieldRangeAccessor<byte>.GetRange(ExcelDataReader reader, long dataOffset, int ordinal, byte[] buffer, int bufferOffset, int length)
	{
		return reader.GetBytes(ordinal, dataOffset, buffer, bufferOffset, length);
	}

	long IFieldRangeAccessor<char>.GetRange(ExcelDataReader reader, long dataOffset, int ordinal, char[] buffer, int bufferOffset, int length)
	{
		return reader.GetChars(ordinal, dataOffset, buffer, bufferOffset, length);
	}

	char IFieldAccessor<char>.GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetChar(ordinal);
	}

	bool IFieldAccessor<bool>.GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetBoolean(ordinal);
	}

	byte IFieldAccessor<byte>.GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetByte(ordinal);
	}

	short IFieldAccessor<short>.GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetInt16(ordinal);
	}

	int IFieldAccessor<int>.GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetInt32(ordinal);
	}

	long IFieldAccessor<long>.GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetInt64(ordinal);
	}

	Guid IFieldAccessor<Guid>.GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetGuid(ordinal);
	}

	float IFieldAccessor<float>.GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetFloat(ordinal);
	}

	double IFieldAccessor<double>.GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetDouble(ordinal);
	}

	DateTime IFieldAccessor<DateTime>.GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetDateTime(ordinal);
	}

	TimeSpan IFieldAccessor<TimeSpan>.GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetTimeSpan(ordinal);
	}

#if DATE_ONLY

	DateOnly IFieldAccessor<DateOnly>.GetValue(ExcelDataReader reader, int ordinal)
	{
		return DateOnlyAccessor.Instance.GetValue(reader, ordinal);
	}

	TimeOnly IFieldAccessor<TimeOnly>.GetValue(ExcelDataReader reader, int ordinal)
	{
		return TimeOnlyAccessor.Instance.GetValue(reader, ordinal);
	}

#endif

	decimal IFieldAccessor<decimal>.GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetDecimal(ordinal);
	}

	byte[] IFieldAccessor<byte[]>.GetValue(ExcelDataReader reader, int ordinal)
	{
		var len = reader.GetBytes(ordinal, 0, null, 0, 0);
		var buffer = new byte[len];
		reader.GetBytes(ordinal, 0, buffer, 0, (int)len);
		return buffer;
	}

	char[] IFieldAccessor<char[]>.GetValue(ExcelDataReader reader, int ordinal)
	{
		var len = reader.GetChars(ordinal, 0, null, 0, 0);
		var buffer = new char[len];
		reader.GetChars(ordinal, 0, buffer, 0, (int)len);
		return buffer;
	}

	Stream IFieldAccessor<Stream>.GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetStream(ordinal);
	}

	TextReader IFieldAccessor<TextReader>.GetValue(ExcelDataReader reader, int ordinal)
	{
		return reader.GetTextReader(ordinal);
	}
}
