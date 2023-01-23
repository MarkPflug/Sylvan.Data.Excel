using System;
using System.Data.Common;
using System.Globalization;
using System.IO;

namespace Sylvan.Data.Excel.Xlsx;

partial class XlsxDataWriter
{
	static DateTime Epoch = new DateTime(1900, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);

	sealed class Context
	{
		public Context(XlsxDataWriter dw, TextWriter xw, DbDataReader dr)
		{
			this.dw = dw;
			this.xw = xw;
			this.dr = dr;
		}

		internal XlsxDataWriter dw;
		internal TextWriter xw;
		internal DbDataReader dr;
		internal char[]? charBuffer;
		internal byte[]? byteBuffer;

		public char[] GetCharBuffer()
		{
			return charBuffer ?? (charBuffer = new char[64]);
		}

		public byte[] GetByteBuffer()
		{
			return byteBuffer ?? (byteBuffer = new byte[48]);
		}
	}

	static class XlsxValueWriter
	{
		static readonly char[] HexMap = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'a', 'b', 'c', 'd', 'e', 'f' };

		const string StringTooLongMessage = "String exceeds the maximum allowed length.";
		static readonly TimeOnly Midnight = new TimeOnly(0);

		public static void WriteString(Context c, string value)
		{
			var w = c.xw;
			w.Write("<c t=\"s\"><v>");
			// truncate before adding to the sharestrings table.
			if (value.Length > StringLimit)
			{
				if (c.dw.truncateStrings)
				{
					value = value.Substring(0, StringLimit);
				}
				else
				{
					throw new FormatException(StringTooLongMessage);
				}
			}

			var ssIdx = c.dw.sharedStrings.GetString(value);
			w.Write(ssIdx);
			w.Write("</v></c>");
		}

		public static void WriteChar(Context c, char value)
		{
			var w = c.xw;
			w.Write("<c t=\"str\"><v>");
			w.Write(value);
			w.Write("</v></c>");
		}

		public static void WriteByte(Context c, byte value)
		{
			var w = c.xw;
			w.Write("<c><v>");

#if SPAN
			var scratch = c.GetCharBuffer();
			if (value.TryFormat(scratch.AsSpan(), out var sl, default, CultureInfo.InvariantCulture))
			{
				w.Write(scratch, 0, sl);
			}
#else
			w.Write(value.ToString(CultureInfo.InvariantCulture));
#endif

			w.Write("</v></c>");
		}

		public static void WriteInt16(Context c, short value)
		{
			var w = c.xw;
			w.Write("<c><v>");

#if SPAN
			var scratch = c.GetCharBuffer();
			if (value.TryFormat(scratch.AsSpan(), out var sl, default, CultureInfo.InvariantCulture))
			{
				w.Write(scratch, 0, sl);
			}
#else
			w.Write(value.ToString(CultureInfo.InvariantCulture));
#endif

			w.Write("</v></c>");
		}

		public static void WriteInt32(Context c, int value)
		{
			var w = c.xw;
			w.Write("<c><v>");

#if SPAN
			var scratch = c.GetCharBuffer();
			if (value.TryFormat(scratch.AsSpan(), out var sl, default, CultureInfo.InvariantCulture))
			{
				w.Write(scratch, 0, sl);
			}
#else
			w.Write(value.ToString(CultureInfo.InvariantCulture));
#endif

			w.Write("</v></c>");
		}

		public static void WriteInt64(Context c, long value)
		{
			var w = c.xw;
			w.Write("<c><v>");

#if SPAN
			var scratch = c.GetCharBuffer();
			if (value.TryFormat(scratch.AsSpan(), out var sl, default, CultureInfo.InvariantCulture))
			{
				w.Write(scratch, 0, sl);
			}
#else
			w.Write(value.ToString(CultureInfo.InvariantCulture));
#endif

			w.Write("</v></c>");
		}

		public static void WriteSingle(Context c, float value)
		{
			var w = c.xw;
			w.Write("<c><v>");

#if SPAN
			var scratch = c.GetCharBuffer();
			if (value.TryFormat(scratch.AsSpan(), out var sl, default, CultureInfo.InvariantCulture))
			{
				w.Write(scratch, 0, sl);
			}
#else
			w.Write(value.ToString(CultureInfo.InvariantCulture));
#endif

			w.Write("</v></c>");
		}

		public static void WriteDouble(Context c, double value)
		{
			var w = c.xw;
			w.Write("<c><v>");

#if SPAN
			var scratch = c.GetCharBuffer();
			if (value.TryFormat(scratch.AsSpan(), out var sl, default, CultureInfo.InvariantCulture))
			{
				w.Write(scratch, 0, sl);
			}
#else
			w.Write(value.ToString(CultureInfo.InvariantCulture));
#endif

			w.Write("</v></c>");
		}

		public static void WriteDecimal(Context c, decimal value)
		{
			var w = c.xw;
			w.Write("<c><v>");

#if SPAN
			var scratch = c.GetCharBuffer();
			if (value.TryFormat(scratch.AsSpan(), out var sl, default, CultureInfo.InvariantCulture))
			{
				w.Write(scratch, 0, sl);
			}
#else
			w.Write(value.ToString(CultureInfo.InvariantCulture));
#endif
			w.Write("</v></c>");
		}

		public static void WriteGuid(Context c, Guid value)
		{
			var w = c.xw;
			// TODO: currently writing these as inline string.
			// might make sense to put in shared string table instead.
			w.Write("<c t=\"str\"><v>");

#if SPAN
			var scratch = c.GetCharBuffer();
			if (value.TryFormat(scratch.AsSpan(), out var sl))
			{
				w.Write(scratch, 0, sl);
			}
#else
			w.Write(value);
#endif

			w.Write("</v></c>");
		}

		public static void WriteBoolean(Context c, bool value)
		{
			var w = c.xw;
			w.Write("<c t=\"b\"><v>");
			w.Write(value ? '1' : '0');
			w.Write("</v></c>");
		}

		public static void WriteDateTime(Context c, DateTime value)
		{
			var w = c.xw;
			w.Write("<c s=\"1\"><v>");

			var val = (value - Epoch).TotalDays + 2;
#if SPAN
			var scratch = c.GetCharBuffer();
			if (val.TryFormat(scratch.AsSpan(), out var sl, default, CultureInfo.InvariantCulture))
			{
				w.Write(scratch, 0, sl);
			}
#else
			w.Write(val.ToString(CultureInfo.InvariantCulture));
#endif
			w.Write("</v></c>");
		}

		public static void WriteTimeSpan(Context c, TimeSpan value)
		{
			var val = value.TotalSeconds;
			var w = c.xw;
			// TODO: currently writing these as inline string.
			// might make sense to put in shared string table instead.
			w.Write("<c><v>");

#if SPAN
			var scratch = c.GetCharBuffer();
			if (val.TryFormat(scratch.AsSpan(), out var sl, default, CultureInfo.InvariantCulture))
			{
				w.Write(scratch, 0, sl);
			}
#else
			w.Write(val.ToString(CultureInfo.InvariantCulture));
#endif

			w.Write("</v></c>");
		}

#if DATE_ONLY

		public static void WriteDateOnly(Context c, DateOnly value)
		{
			var w = c.xw;
			w.Write("<c s=\"2\"><v>");

			var val = (value.ToDateTime(Midnight) - Epoch).TotalDays + 2;

			var scratch = c.GetCharBuffer();
			if (val.TryFormat(scratch.AsSpan(), out var sl, default, CultureInfo.InvariantCulture))
			{
				w.Write(scratch, 0, sl);
			}

			w.Write("</v></c>");
		}

		public static void WriteTimeOnly(Context c, TimeOnly value)
		{
			var w = c.xw;
			w.Write("<c s=\"3\"><v>");

			var val = value.ToTimeSpan().TotalDays;

			var scratch = c.GetCharBuffer();
			if (val.TryFormat(scratch.AsSpan(), out var sl, default, CultureInfo.InvariantCulture))
			{
				w.Write(scratch, 0, sl);
			}

			w.Write("</v></c>");
		}
#endif

		public static void WriteBinaryHex(Context c, Func<byte[], int> reader)
		{
			var w = c.xw;
			w.Write("<c t=\"str\"><v>");
			var idx = 0;
			var buffer = c.GetByteBuffer();
			var charBuffer = c.GetCharBuffer();
			int len;
			var pos = 0;
			
			w.Write("0x");
			while ((len = reader(buffer)) != 0)
			{
				var l = ToHexCharArray(buffer, 0, len, charBuffer, pos);
				w.Write(charBuffer, 0, l);
				idx += len;
				pos += l;
			}

			w.Write("</v></c>");

		}

		static int ToHexCharArray(byte[] dataBuffer, int offset, int length, char[] outputBuffer, int outputOffset)
		{
			if (length * 2 > outputBuffer.Length - outputOffset)
				throw new ArgumentException();

			var idx = offset;
			var end = offset + length;
			for (; idx < end; idx++)
			{
				var b = dataBuffer[idx];
				var lo = HexMap[b & 0xf];
				var hi = HexMap[b >> 4];
				outputBuffer[outputOffset++] = hi;
				outputBuffer[outputOffset++] = lo;
			}
			return length * 2;
		}

		static void WriteCharArray(Context c, Func<char[],int> reader)
		{
			var w = c.xw;
			var buffer = c.GetCharBuffer();
			w.Write("<c t=\"str\"><v>");
			var idx = 0;
			int len;
			
			while ((len = reader(buffer)) != 0)
			{
				w.Write(buffer, 0, len);
				idx += len;
			}

			w.Write("</v></c>");
		}
	}

	abstract class FieldWriter
	{
		public static FieldWriter Get(Type type)
		{
			var code = Type.GetTypeCode(type);

			switch (code)
			{
				case TypeCode.Boolean:
					return new BooleanFieldWriter();
				case TypeCode.Char:
					return new CharFieldWriter();
				case TypeCode.DateTime:
					return new DateTimeFieldWriter();
				case TypeCode.String:
					return new StringFieldWriter();
				case TypeCode.Byte:
					return new ByteFieldWriter();
				case TypeCode.Int16:
					return new Int16FieldWriter();
				case TypeCode.Int32:
					return new Int32FieldWriter();
				case TypeCode.Int64:
					return new Int64FieldWriter();
				case TypeCode.Single:
					return new SingleFieldWriter();
				case TypeCode.Double:
					return new DoubleFieldWriter();
				case TypeCode.Decimal:
					return new DecimalFieldWriter();
				default:
					if (type == typeof(byte[]))
					{
						return new BinaryHexFieldWriter();
					}
					if (type == typeof(char[]))
					{
						return new CharArrayFieldWriter();
					}
					if (type == typeof(Guid))
					{
						return new GuidFieldWriter();
					}
					if (type == typeof(TimeSpan))
					{
						return new TimeSpanFieldWriter();
					}

#if DATE_ONLY
					if (type == typeof(DateOnly))
					{
						return new DateOnlyFieldWriter();
					}

					if (type == typeof(TimeOnly))
					{
						return new TimeOnlyFieldWriter();
					}
#endif

					return new ObjectFieldWriter();
			}
		}

		public abstract void WriteField(Context c, int ordinal);

		public virtual double GetWidth(DbDataReader data, int ordinal)
		{
			return 12;
		}
	}

	sealed class ObjectFieldWriter : FieldWriter
	{
		public static readonly ObjectFieldWriter Instance = new ObjectFieldWriter();

		public override void WriteField(Context c, int ordinal)
		{
			var val = c.dr.GetValue(ordinal);

			var type = val.GetType();

			var tc = Type.GetTypeCode(type);

			switch (tc)
			{
				case TypeCode.String:
					XlsxValueWriter.WriteString(c, (string)val);
					break;
				default:
					var str = val?.ToString() ?? string.Empty;
					XlsxValueWriter.WriteString(c, str);
					break;
			}
		}
	}

	sealed class StringFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var value = c.dr.GetString(ordinal);
			XlsxValueWriter.WriteString(c, value);
		}
	}

	sealed class BooleanFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var val = c.dr.GetBoolean(ordinal);
			XlsxValueWriter.WriteBoolean(c, val);
		}
	}

	sealed class DateTimeFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var dt = c.dr.GetDateTime(ordinal);
			XlsxValueWriter.WriteDateTime(c, dt);			
		}

		public override double GetWidth(DbDataReader data, int ordinal)
		{
			return 22;
		}
	}

#if DATE_ONLY

	sealed class DateOnlyFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var dt = c.dr.GetFieldValue<DateOnly>(ordinal);
			XlsxValueWriter.WriteDateOnly(c, dt);
		}

		public override double GetWidth(DbDataReader data, int ordinal)
		{
			return 11;
		}
	}

	sealed class TimeOnlyFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var value = c.dr.GetFieldValue<TimeOnly>(ordinal);
			XlsxValueWriter.WriteTimeOnly(c, value);
		}

		public override double GetWidth(DbDataReader data, int ordinal)
		{
			return 11;
		}
	}

#endif

	sealed class DecimalFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var val = c.dr.GetDecimal(ordinal);
			XlsxValueWriter.WriteDecimal(c, val);			
		}
	}

	sealed class SingleFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var val = c.dr.GetFloat(ordinal);
			XlsxValueWriter.WriteSingle(c, val);
		}
	}

	sealed class DoubleFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var val = c.dr.GetDouble(ordinal);
			XlsxValueWriter.WriteDouble(c, val);
		}
	}

	sealed class CharFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var val = c.dr.GetChar(ordinal);
			XlsxValueWriter.WriteChar(c, val);
		}
	}

	sealed class ByteFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var val = c.dr.GetByte(ordinal);
			XlsxValueWriter.WriteByte(c, val);
		}
	}

	sealed class Int16FieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var val = c.dr.GetInt16(ordinal);
			XlsxValueWriter.WriteInt16(c, val);
		}
	}

	sealed class Int32FieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var val = c.dr.GetInt32(ordinal);
			XlsxValueWriter.WriteInt32(c, val);
		}
	}

	sealed class Int64FieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var val = c.dr.GetInt64(ordinal);
			XlsxValueWriter.WriteInt64(c, val);
		}
	}

	sealed class GuidFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var val = c.dr.GetGuid(ordinal);
			XlsxValueWriter.WriteGuid(c, val);
		}

		public override double GetWidth(DbDataReader data, int ordinal)
		{
			return 38;
		}
	}

	sealed class TimeSpanFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var val = c.dr.GetFieldValue<TimeSpan>(ordinal);
			XlsxValueWriter.WriteTimeSpan(c, val);	
		}
	}

	sealed class BinaryHexFieldWriter : FieldWriter
	{

		byte[] dataBuffer = new byte[48];

		public override void WriteField(Context context, int ordinal)
		{
			XlsxValueWriter.write
		}		
	}

	sealed class CharArrayFieldWriter : FieldWriter
	{
		public override void WriteField(Context context, int ordinal)
		{
			
			
		}
	}
}
