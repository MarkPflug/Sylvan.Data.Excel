#if NET6_0_OR_GREATER

using System;
using System.Data.Common;
using System.IO;

namespace Sylvan.Data.Excel.Xlsb;

partial class XlsbDataWriter
{
	static DateTime Epoch = new DateTime(1900, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);

	sealed class Context
	{
		public Context(XlsbDataWriter dw, BinaryWriter bw, DbDataReader dr)
		{
			this.dw = dw;
			this.bw = bw;
			this.dr = dr;
		}

		internal XlsbDataWriter dw;
		internal BinaryWriter bw;
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

	static class XlsbValueWriter
	{
		static readonly char[] HexMap = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'a', 'b', 'c', 'd', 'e', 'f' };

		const string StringTooLongMessage = "String exceeds the maximum allowed length.";


		public static void WriteString(Context c, int col, string value)
		{
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
			var ssidx = c.dw.sharedStrings.GetString(value);
			c.bw.WriteSharedString(col, ssidx);
		}

		public static void WriteChar(Context c, int col, char value)
		{
			var w = c.bw;
			var idx = c.dw.sharedStrings.GetString(value.ToString());
			w.WriteSharedString(col, idx);
		}

		public static void WriteByte(Context c, int col, byte value)
		{
			var w = c.bw;
			w.WriteNumber(col, value);
		}

		public static void WriteInt16(Context c, int col, short value)
		{
			var w = c.bw;
			w.WriteNumber(col, value);
		}

		public static void WriteInt32(Context c, int col, int value)
		{
			var w = c.bw;
			w.WriteNumber(col, value);
		}

		public static void WriteInt64(Context c, int col, long value)
		{
			var w = c.bw;
			if (value == (int)value)
			{
				w.WriteNumber(col, (int)value);
			}
			else
			{
				w.WriteNumber(col, (double)value);
			}
		}

		public static void WriteSingle(Context c, int col, float value)
		{

			var w = c.bw;
			w.WriteNumber(col, value);
		}

		public static void WriteDouble(Context c, int col, double value)
		{
			var w = c.bw;
			w.WriteNumber(col, value);
		}

		public static void WriteDecimal(Context c, int col, decimal value)
		{
			var w = c.bw;
			w.WriteNumber(col, value);
		}

		public static void WriteGuid(Context c, int col, Guid value)
		{
			var w = c.bw;
			var str = value.ToString();
			var idx = c.dw.sharedStrings.GetString(str);
			w.WriteSharedString(col, idx);

		}

		public static void WriteBoolean(Context c, int col, bool value)
		{
			c.bw.WriteBool(col, value);
		}

		public static void WriteDateTime(Context c, int col, DateTime value)
		{
			var val = (value - Epoch).TotalDays + 2;
			//TODO: format
			c.bw.WriteNumber(col, val);
		}

		public static void WriteTimeSpan(Context c, int col, TimeSpan value)
		{
			var val = value.TotalDays;
			//TODO: format
			c.bw.WriteNumber(col, val);
		}

#if DATE_ONLY

		static readonly TimeOnly Midnight = new TimeOnly(0);

		public static void WriteDateOnly(Context c, int col, DateOnly value)
		{
			var w = c.bw;
			var val = (value.ToDateTime(Midnight) - Epoch).TotalDays + 2;
			w.WriteNumber(col, val);
		}

		public static void WriteTimeOnly(Context c, int col, TimeOnly value)
		{
			var w = c.bw;
			var val = value.ToTimeSpan().TotalDays;
			w.WriteNumber(col, val);
		}
#endif

		public static void WriteBinaryHex(Context c, int col, byte[] value)
		{
			throw new NotImplementedException();
			//var w = c.bw;
			//w.Write("<c t=\"str\"><v>");
			//var charBuffer = c.GetCharBuffer();
			//var idx = 0;
			//w.Write("0x");
			//while (idx < value.Length)
			//{
			//	var l = ToHexCharArray(value, idx, 48, charBuffer, 0);
			//	w.Write(charBuffer, 0, l);
			//	idx += 48;
			//}

			//w.Write("</v></c>");
		}

		public static int ToHexCharArray(byte[] dataBuffer, int offset, int length, char[] outputBuffer, int outputOffset)
		{
			throw new NotImplementedException();
			//if (length * 2 > outputBuffer.Length - outputOffset)
			//	throw new ArgumentException();

			//var idx = offset;
			//var end = offset + length;
			//for (; idx < end; idx++)
			//{
			//	var b = dataBuffer[idx];
			//	var lo = HexMap[b & 0xf];
			//	var hi = HexMap[b >> 4];
			//	outputBuffer[outputOffset++] = hi;
			//	outputBuffer[outputOffset++] = lo;
			//}
			//return length * 2;
		}

		public static void WriteCharArray(Context c, int col, char[] value)
		{
			throw new NotImplementedException();
			//var w = c.bw;
			//w.Write("<c t=\"str\"><v>");
			//// TODO: limit length...
			//w.Write(value);
			//w.Write("</v></c>");
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
					XlsbValueWriter.WriteBoolean(c, ordinal, (bool)val);
					break;
				case TypeCode.String:
					XlsbValueWriter.WriteString(c, ordinal, (string)val);
					break;
				case TypeCode.Byte:
					XlsbValueWriter.WriteByte(c, ordinal, (byte)val);
					break;
				case TypeCode.Int16:
					XlsbValueWriter.WriteInt16(c, ordinal, (short)val);
					break;
				case TypeCode.Int32:
					XlsbValueWriter.WriteInt32(c, ordinal, (int)val);
					break;
				case TypeCode.Int64:
					XlsbValueWriter.WriteInt64(c, ordinal, (long)val);
					break;
				case TypeCode.DateTime:
					XlsbValueWriter.WriteDateTime(c, ordinal, (DateTime)val);
					break;
				case TypeCode.Single:
					XlsbValueWriter.WriteSingle(c, ordinal, (float)val);
					break;
				case TypeCode.Double:
					XlsbValueWriter.WriteDouble(c, ordinal, (double)val);
					break;
				case TypeCode.Decimal:
					XlsbValueWriter.WriteDecimal(c, ordinal, (decimal)val);
					break;
				default:

					if (type == typeof(byte[]))
					{
						XlsbValueWriter.WriteBinaryHex(c, ordinal, (byte[])val);
						break;
					}
					if (type == typeof(char[]))
					{
						XlsbValueWriter.WriteCharArray(c, ordinal, (char[])val);
						break;
					}
					if (type == typeof(Guid))
					{
						XlsbValueWriter.WriteGuid(c, ordinal, (Guid)val);
						break;
					}
					if (type == typeof(TimeSpan))
					{
						XlsbValueWriter.WriteTimeSpan(c, ordinal, (TimeSpan)val);
						break;
					}

#if DATE_ONLY
					if (type == typeof(DateOnly))
					{
						XlsbValueWriter.WriteDateOnly(c, ordinal, (DateOnly)val);
						break;
					}

					if (type == typeof(TimeOnly))
					{
						XlsbValueWriter.WriteTimeOnly(c, ordinal, (TimeOnly)val);
						break;
					}
#endif
					// anything else, we'll just ToString
					var str = val?.ToString() ?? string.Empty;
					XlsbValueWriter.WriteString(c, ordinal, str);
					break;
			}
		}
	}

	sealed class StringFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var value = c.dr.GetString(ordinal);
			XlsbValueWriter.WriteString(c, ordinal, value);
		}
	}

	sealed class BooleanFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var val = c.dr.GetBoolean(ordinal);
			XlsbValueWriter.WriteBoolean(c, ordinal, val);
		}
	}

	sealed class DateTimeFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var dt = c.dr.GetDateTime(ordinal);
			XlsbValueWriter.WriteDateTime(c, ordinal, dt);
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
			XlsbValueWriter.WriteDateOnly(c, ordinal, dt);
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
			XlsbValueWriter.WriteTimeOnly(c, ordinal, value);
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
			XlsbValueWriter.WriteDecimal(c, ordinal, val);
		}
	}

	sealed class SingleFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var val = c.dr.GetFloat(ordinal);
			XlsbValueWriter.WriteSingle(c, ordinal, val);
		}
	}

	sealed class DoubleFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var val = c.dr.GetDouble(ordinal);
			XlsbValueWriter.WriteDouble(c, ordinal, val);
		}
	}

	sealed class CharFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var val = c.dr.GetChar(ordinal);
			XlsbValueWriter.WriteChar(c, ordinal, val);
		}
	}

	sealed class ByteFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var val = c.dr.GetByte(ordinal);
			XlsbValueWriter.WriteByte(c, ordinal, val);
		}
	}

	sealed class Int16FieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var val = c.dr.GetInt16(ordinal);
			XlsbValueWriter.WriteInt16(c, ordinal, val);
		}
	}

	sealed class Int32FieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var val = c.dr.GetInt32(ordinal);
			XlsbValueWriter.WriteInt32(c, ordinal, val);
		}
	}

	sealed class Int64FieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var val = c.dr.GetInt64(ordinal);
			XlsbValueWriter.WriteInt64(c, ordinal, val);
		}
	}

	sealed class GuidFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var val = c.dr.GetGuid(ordinal);
			XlsbValueWriter.WriteGuid(c, ordinal, val);
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
			XlsbValueWriter.WriteTimeSpan(c, ordinal, val);
		}
	}

	sealed class BinaryHexFieldWriter : FieldWriter
	{
		public override void WriteField(Context context, int ordinal)
		{
			throw new NotImplementedException();
			//var w = context.bw;
			//w.Write("<c t=\"str\"><v>");
			//var idx = 0;
			//var dataBuffer = context.GetByteBuffer();
			//var charBuffer = context.GetCharBuffer();
			//int len;
			//var reader = context.dr;
			//w.Write("0x");
			//while ((len = (int)reader.GetBytes(ordinal, idx, dataBuffer, 0, dataBuffer.Length)) != 0)
			//{
			//	var c = XlsbValueWriter.ToHexCharArray(dataBuffer, 0, len, charBuffer, 0);
			//	w.Write(charBuffer, 0, c);
			//	idx += len;
			//}

			//w.Write("</v></c>");
		}
	}

	sealed class CharArrayFieldWriter : FieldWriter
	{
		public override void WriteField(Context context, int ordinal)
		{

			throw new NotImplementedException();
			//var w = context.bw;
			//w.Write("<c t=\"str\"><v>");
			//var idx = 0;
			//var dataBuffer = context.GetCharBuffer();
			//int len;
			//var reader = context.dr;
			//while ((len = (int)reader.GetChars(ordinal, idx, dataBuffer, 0, dataBuffer.Length)) != 0)
			//{
			//	w.Write(dataBuffer, 0, len);
			//	idx += len;
			//}

			//w.Write("</v></c>");
		}
	}

}

#endif