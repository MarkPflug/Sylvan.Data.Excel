﻿using System;
using System.Data.Common;
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
	}

	abstract class FieldWriter
	{
		protected char[] scratch = new char[100];

		public static FieldWriter Get(Type type)
		{
			var code = Type.GetTypeCode(type);

			switch (code)
			{
				case TypeCode.Boolean:
					return new BooleanFieldWriter();
				case TypeCode.DateTime:
					return new DateTimeFieldWriter();
				case TypeCode.String:
					return new StringFieldWriter();
				case TypeCode.Int32:
					return new Int32FieldWriter();
				case TypeCode.Int64:
					return new Int64FieldWriter();
				case TypeCode.Double:
					return new DoubleFieldWriter();
				case TypeCode.Decimal:
					return new DecimalFieldWriter();
				default:
#if DATE_ONLY
					if (type == typeof(DateOnly))
					{
						return new DateOnlyFieldWriter();
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
		public override void WriteField(Context c, int ordinal)
		{
			var w = c.xw;
			w.Write("<c t=\"s\"><v>");
			var s = c.dr.GetString(ordinal);
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

			if (val.TryFormat(scratch.AsSpan(), out var sl))
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
			if (val.TryFormat(scratch.AsSpan(), out var sl))
			{
				w.Write(scratch, 0, sl);
			}
#else
			w.Write(val);
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
			if (val.TryFormat(scratch.AsSpan(), out var sl))
			{
				w.Write(scratch, 0, sl);
			}
#else
			w.Write(val);
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
			if (val.TryFormat(scratch.AsSpan(), out var sl))
			{
				w.Write(scratch, 0, sl);
			}
#else
			w.Write(val);
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
			if (val.TryFormat(scratch.AsSpan(), out var sl))
			{
				w.Write(scratch, 0, sl);
			}
#else
			w.Write(val);
#endif

			w.Write("</v></c>");
		}
	}
}
