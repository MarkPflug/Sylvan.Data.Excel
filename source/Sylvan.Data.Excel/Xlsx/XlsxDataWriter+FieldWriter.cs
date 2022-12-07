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
		internal char[]? scratch;

		public char[] GetScratch()
		{
			return scratch ?? (scratch = new char[64]);
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
			var w = c.xw;
			w.Write("<c t=\"s\"><v>");

			var s = val?.ToString() ?? "";
			var ssIdx = c.dw.sharedStrings.GetString(s);
			w.Write(ssIdx);
			w.Write("</v></c>");
		}
	}

	sealed class StringFieldWriter : FieldWriter
	{
		const string StringTooLongMessage = "String exceeds the maximum allowed length.";

		public override void WriteField(Context c, int ordinal)
		{
			var w = c.xw;
			w.Write("<c t=\"s\"><v>");
			var s = c.dr.GetString(ordinal);
			// truncate before adding to the sharestrings table.
			if (s.Length > StringLimit)
			{
				if (c.dw.truncateStrings)
				{
					s = s.Substring(0, StringLimit);
				}
				else
				{
					throw new FormatException(StringTooLongMessage);
				}
			}

			var ssIdx = c.dw.sharedStrings.GetString(s);
			w.Write(ssIdx);
			w.Write("</v></c>");
		}
	}

	sealed class BooleanFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var w = c.xw;
			w.Write("<c t=\"b\"><v>");
			var val = c.dr.GetBoolean(ordinal);
			w.Write(val ? '1' : '0');
			w.Write("</v></c>");
		}
	}

	sealed class DateTimeFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var w = c.xw;
			w.Write("<c s=\"1\"><v>");

			var dt = c.dr.GetDateTime(ordinal);
			var val = (dt - Epoch).TotalDays + 2;
#if SPAN
			var scratch = c.GetScratch();
			if (val.TryFormat(scratch.AsSpan(), out var sl, default, CultureInfo.InvariantCulture))
			{
				w.Write(scratch, 0, sl);
			}
#else
			w.Write(val.ToString(CultureInfo.InvariantCulture));
#endif
			w.Write("</v></c>");
		}

		public override double GetWidth(DbDataReader data, int ordinal)
		{
			return 22;
		}
	}

#if DATE_ONLY

	sealed class DateOnlyFieldWriter : FieldWriter
	{
		static readonly TimeOnly Midnight = new TimeOnly(0);

		public override void WriteField(Context c, int ordinal)
		{
			var w = c.xw;
			w.Write("<c s=\"2\"><v>");

			var dt = c.dr.GetFieldValue<DateOnly>(ordinal);
			var val = (dt.ToDateTime(Midnight) - Epoch).TotalDays + 2;

			var scratch = c.GetScratch();
			if (val.TryFormat(scratch.AsSpan(), out var sl, default, CultureInfo.InvariantCulture))
			{
				w.Write(scratch, 0, sl);
			}

			w.Write("</v></c>");
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
			var w = c.xw;
			w.Write("<c s=\"3\"><v>");

			var t = c.dr.GetFieldValue<TimeOnly>(ordinal);
			var val = t.ToTimeSpan().TotalDays;

			var scratch = c.GetScratch();
			if (val.TryFormat(scratch.AsSpan(), out var sl, default, CultureInfo.InvariantCulture))
			{
				w.Write(scratch, 0, sl);
			}

			w.Write("</v></c>");
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
			var w = c.xw;
			w.Write("<c><v>");

			var val = c.dr.GetDecimal(ordinal);
#if SPAN
			var scratch = c.GetScratch();
			if (val.TryFormat(scratch.AsSpan(), out var sl, default, CultureInfo.InvariantCulture))
			{
				w.Write(scratch, 0, sl);
			}
#else
			w.Write(val.ToString(CultureInfo.InvariantCulture));
#endif

			w.Write("</v></c>");
		}
	}

	sealed class SingleFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var w = c.xw;
			w.Write("<c><v>");

			var val = c.dr.GetFloat(ordinal);
#if SPAN
			var scratch = c.GetScratch();
			if (val.TryFormat(scratch.AsSpan(), out var sl, default, CultureInfo.InvariantCulture))
			{
				w.Write(scratch, 0, sl);
			}
#else
			w.Write(val.ToString(CultureInfo.InvariantCulture));
#endif

			w.Write("</v></c>");
		}
	}

	sealed class DoubleFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var w = c.xw;
			w.Write("<c><v>");

			var val = c.dr.GetDouble(ordinal);
#if SPAN
			var scratch = c.GetScratch();
			if (val.TryFormat(scratch.AsSpan(), out var sl, default, CultureInfo.InvariantCulture))
			{
				w.Write(scratch, 0, sl);
			}
#else
			w.Write(val.ToString(CultureInfo.InvariantCulture));
#endif

			w.Write("</v></c>");
		}
	}

	sealed class CharFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var w = c.xw;
			w.Write("<c t=\"str\"><v>");

			var val = c.dr.GetChar(ordinal);
			w.Write(val);
			w.Write("</v></c>");
		}
	}

	sealed class ByteFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var w = c.xw;
			w.Write("<c><v>");

			var val = c.dr.GetByte(ordinal);
#if SPAN
			var scratch = c.GetScratch();
			if (val.TryFormat(scratch.AsSpan(), out var sl, default, CultureInfo.InvariantCulture))
			{
				w.Write(scratch, 0, sl);
			}
#else
			w.Write(val.ToString(CultureInfo.InvariantCulture));
#endif

			w.Write("</v></c>");
		}
	}

	sealed class Int16FieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var w = c.xw;
			w.Write("<c><v>");

			var val = c.dr.GetInt16(ordinal);
#if SPAN
			var scratch = c.GetScratch();
			if (val.TryFormat(scratch.AsSpan(), out var sl, default, CultureInfo.InvariantCulture))
			{
				w.Write(scratch, 0, sl);
			}
#else
			w.Write(val.ToString(CultureInfo.InvariantCulture));
#endif

			w.Write("</v></c>");
		}
	}

	sealed class Int32FieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var w = c.xw;
			w.Write("<c><v>");

			var val = c.dr.GetInt32(ordinal);
#if SPAN
			var scratch = c.GetScratch();
			if (val.TryFormat(scratch.AsSpan(), out var sl, default, CultureInfo.InvariantCulture))
			{
				w.Write(scratch, 0, sl);
			}
#else
			w.Write(val.ToString(CultureInfo.InvariantCulture));
#endif

			w.Write("</v></c>");
		}
	}

	sealed class Int64FieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var w = c.xw;
			w.Write("<c><v>");

			var val = c.dr.GetInt64(ordinal);
#if SPAN
			var scratch = c.GetScratch();
			if (val.TryFormat(scratch.AsSpan(), out var sl, default, CultureInfo.InvariantCulture))
			{
				w.Write(scratch, 0, sl);
			}
#else
			w.Write(val.ToString(CultureInfo.InvariantCulture));
#endif

			w.Write("</v></c>");
		}
	}

	sealed class GuidFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var w = c.xw;
			// TODO: currently writing these as inline string.
			// might make sense to put in shared string table instead.
			w.Write("<c t=\"str\"><v>");

			var val = c.dr.GetGuid(ordinal);
#if SPAN
			var scratch = c.GetScratch();
			if (val.TryFormat(scratch.AsSpan(), out var sl))
			{
				w.Write(scratch, 0, sl);
			}
#else
			w.Write(val);
#endif

			w.Write("</v></c>");
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
			var w = c.xw;
			// TODO: currently writing these as inline string.
			// might make sense to put in shared string table instead.
			w.Write("<c><v>");

			var val = c.dr.GetFieldValue<TimeSpan>(ordinal).TotalSeconds;
#if SPAN
			var scratch = c.GetScratch();
			if (val.TryFormat(scratch.AsSpan(), out var sl, default, CultureInfo.InvariantCulture))
			{
				w.Write(scratch, 0, sl);
			}
#else
			w.Write(val.ToString(CultureInfo.InvariantCulture));
#endif

			w.Write("</v></c>");
		}
	}

	sealed class BinaryHexFieldWriter : FieldWriter
	{
		static char[] HexMap = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'a', 'b', 'c', 'd', 'e', 'f' };

		byte[] dataBuffer = new byte[48];

		public override void WriteField(Context context, int ordinal)
		{
			var w = context.xw;
			w.Write("<c t=\"str\"><v>");
			var idx = 0;
			var buffer = context.GetScratch();
			int len;
			var pos = 0;
			var reader = context.dr;
			w.Write("0x");
			while ((len = (int)reader.GetBytes(ordinal, idx, dataBuffer, 0, dataBuffer.Length)) != 0)
			{
				var c = ToHexCharArray(dataBuffer, 0, len, buffer, pos);
				w.Write(buffer, 0, c);
				idx += len;
				pos += c;
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
	}

	sealed class CharArrayFieldWriter : FieldWriter
	{
		char[] dataBuffer = new char[128];

		public override void WriteField(Context context, int ordinal)
		{
			var w = context.xw;
			w.Write("<c t=\"str\"><v>");
			var idx = 0;
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
