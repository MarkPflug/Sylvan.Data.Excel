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
		const string StringTooLongMessage = "String exceeds the maximum allowed length.";
		

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
			var val = value.TotalDays;
			var w = c.xw;
			// TODO: currently writing these as inline string.
			// might make sense to put in shared string table instead.
			w.Write("<c s=\"3\"><v>");

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

		static readonly TimeOnly Midnight = new TimeOnly(0);

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

		public static void WriteBinaryHex(Context c, byte[] value)
		{
			var w = c.xw;
			w.Write("<c t=\"str\"><v>");
			var charBuffer = c.GetCharBuffer();
			var idx = 0;
			w.Write("0x");
			while (idx < value.Length)
			{
				var l = HexCodec.ToHexCharArray(value, idx, 48, charBuffer, 0);
				w.Write(charBuffer, 0, l);
				idx += 48;				
			}

			w.Write("</v></c>");
		}

		public static void WriteCharArray(Context c, char[] value)
		{
			var w = c.xw;
			w.Write("<c t=\"str\"><v>");			
			// TODO: limit length...
			w.Write(value);
			w.Write("</v></c>");
		}
	}

	abstract class FieldWriter
	{
		public static readonly ObjectFieldWriter Object = new ObjectFieldWriter();
		public static readonly BooleanFieldWriter Boolean = new BooleanFieldWriter();
		public static readonly CharFieldWriter Char = new CharFieldWriter();
		public static readonly StringFieldWriter String = new StringFieldWriter();
		public static readonly ByteFieldWriter Byte = new ByteFieldWriter();
		public static readonly Int16FieldWriter Int16 = new Int16FieldWriter();
		public static readonly Int32FieldWriter Int32 = new Int32FieldWriter();
		public static readonly Int64FieldWriter Int64 = new Int64FieldWriter();
		public static readonly SingleFieldWriter Single = new SingleFieldWriter();
		public static readonly DoubleFieldWriter Double = new DoubleFieldWriter();
		public static readonly DecimalFieldWriter Decimal = new DecimalFieldWriter();
		public static readonly DateTimeFieldWriter DateTime = new DateTimeFieldWriter();
		public static readonly TimeSpanFieldWriter TimeSpan = new TimeSpanFieldWriter();

#if DATE_ONLY
		public static readonly DateOnlyFieldWriter DateOnly = new DateOnlyFieldWriter();
		public static readonly TimeOnlyFieldWriter TimeOnly = new TimeOnlyFieldWriter();
#endif

		public static readonly GuidFieldWriter Guid = new GuidFieldWriter();

		public static FieldWriter Get(Type type)
		{
			var code = Type.GetTypeCode(type);

			switch (code)
			{
				case TypeCode.Boolean:
					return Boolean;
				case TypeCode.Char:
					return Char;
				case TypeCode.DateTime:
					return DateTime;
				case TypeCode.String:
					return String;
				case TypeCode.Byte:
					return Byte;
				case TypeCode.Int16:
					return Int16;
				case TypeCode.Int32:
					return Int32;
				case TypeCode.Int64:
					return Int64;
				case TypeCode.Single:
					return Single;
				case TypeCode.Double:
					return Double;
				case TypeCode.Decimal:
					return Decimal;
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
						return Guid;
					}
					if (type == typeof(TimeSpan))
					{
						return TimeSpan;
					}

#if DATE_ONLY
					if (type == typeof(DateOnly))
					{
						return DateOnly;
					}

					if (type == typeof(TimeOnly))
					{
						return TimeOnly;
					}
#endif

					return Object;
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
		public override void WriteField(Context c, int ordinal)
		{
			var val = c.dr.GetValue(ordinal);

			var type = val.GetType();

			var tc = Type.GetTypeCode(type);

			switch (tc)
			{
				case TypeCode.Boolean:
					XlsxValueWriter.WriteBoolean(c, (bool)val);
					break;
				case TypeCode.String:
					XlsxValueWriter.WriteString(c, (string)val);
					break;
				case TypeCode.Byte:
					XlsxValueWriter.WriteByte(c, (byte)val);
					break;
				case TypeCode.Int16:
					XlsxValueWriter.WriteInt16(c, (short)val);
					break;
				case TypeCode.Int32:
					XlsxValueWriter.WriteInt32(c, (int)val);
					break;
				case TypeCode.Int64:
					XlsxValueWriter.WriteInt64(c, (long)val);
					break;
				case TypeCode.DateTime:
					XlsxValueWriter.WriteDateTime(c, (DateTime)val);
					break;
				case TypeCode.Single:
					XlsxValueWriter.WriteSingle(c, (float)val);
					break;
				case TypeCode.Double:
					XlsxValueWriter.WriteDouble(c, (double)val);
					break;
				case TypeCode.Decimal:
					XlsxValueWriter.WriteDecimal(c, (decimal)val);
					break;
				default:

					if (type == typeof(byte[]))
					{
						XlsxValueWriter.WriteBinaryHex(c, (byte[])val);
						break;
					}
					if (type == typeof(char[]))
					{
						XlsxValueWriter.WriteCharArray(c, (char[])val);
						break;
					}
					if (type == typeof(Guid))
					{
						XlsxValueWriter.WriteGuid(c, (Guid)val);
						break;
					}
					if (type == typeof(TimeSpan))
					{
						XlsxValueWriter.WriteTimeSpan(c, (TimeSpan)val);
						break;
					}

#if DATE_ONLY
					if (type == typeof(DateOnly))
					{
						XlsxValueWriter.WriteDateOnly(c, (DateOnly)val);
						break;
					}

					if (type == typeof(TimeOnly))
					{
						XlsxValueWriter.WriteTimeOnly(c, (TimeOnly)val);
						break;
					}
#endif
					// anything else, we'll just ToString
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
		public override void WriteField(Context context, int ordinal)
		{
			var w = context.xw;
			w.Write("<c t=\"str\"><v>");
			var idx = 0;
			var dataBuffer = context.GetByteBuffer();
			var charBuffer = context.GetCharBuffer();
			int len;
			var reader = context.dr;
			w.Write("0x");
			while ((len = (int)reader.GetBytes(ordinal, idx, dataBuffer, 0, dataBuffer.Length)) != 0)
			{
				var c = HexCodec.ToHexCharArray(dataBuffer, 0, len, charBuffer, 0);
				w.Write(charBuffer, 0, c);
				idx += len;
			}

			w.Write("</v></c>");
		}
	}

	sealed class CharArrayFieldWriter : FieldWriter
	{
		public override void WriteField(Context context, int ordinal)
		{
			var w = context.xw;
			w.Write("<c t=\"str\"><v>");
			var idx = 0;
			var dataBuffer = context.GetCharBuffer();
			int len;
			var reader = context.dr;
			while ((len = (int)reader.GetChars(ordinal, idx, dataBuffer, 0, dataBuffer.Length)) != 0)
			{
				w.Write(dataBuffer, 0, len);
				idx += len;
			}

			w.Write("</v></c>");
		}
	}
}
