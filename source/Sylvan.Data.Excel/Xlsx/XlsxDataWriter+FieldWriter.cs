using System;
using System.Data.Common;
using System.Xml;

namespace Sylvan.Data.Excel.Xlsx;

partial class XlsxDataWriter
{
	sealed class Context
	{
		public Context(XlsxDataWriter dw, XmlWriter xw, DbDataReader dr)
		{
			this.dw = dw;
			this.xw = xw;
			this.dr = dr;
		}

		internal XlsxDataWriter dw;
		internal XmlWriter xw;
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
				case TypeCode.DateTime:
					return new DateTimeFieldWriter();
				case TypeCode.String:
					return new StringFieldWriter();
				case TypeCode.Int32:
					return new Int32FieldWriter();
				case TypeCode.Double:
					return new DoubleFieldWriter();
				case TypeCode.Decimal:
					return new DecimalFieldWriter();
				default:
					throw new NotSupportedException();
			}
		}

		public abstract void WriteField(Context c, int ordinal);
	}

	sealed class StringFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var w = c.xw;
			w.WriteStartElement("c", NS);
			w.WriteStartAttribute("t");
			w.WriteValue("s");
			w.WriteEndAttribute();

			w.WriteStartElement("v", NS);

			var s = c.dr.GetString(ordinal);
			var ssIdx = c.dw.sharedStrings.GetString(s);
			w.WriteValue(ssIdx);

			w.WriteEndElement();
			w.WriteEndElement();
		}
	}

	sealed class DateTimeFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var w = c.xw;
			w.WriteStartElement("c", NS);

			w.WriteStartAttribute("t");
			w.WriteValue("d");
			w.WriteEndAttribute();
			w.WriteStartAttribute("s");

			var fmtId = "2";
			w.WriteValue(fmtId);
			w.WriteEndAttribute();

			w.WriteStartElement("v", NS);

			var dt = c.dr.GetDateTime(ordinal);
			if (IsoDate.TryFormatIso(dt, scratch.AsSpan(), out var sl))
			{
				c.xw.WriteRaw(scratch, 0, sl);
			}

			w.WriteEndElement();
			w.WriteEndElement();
		}
	}

	sealed class DecimalFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var w = c.xw;
			w.WriteStartElement("c", NS);

			w.WriteStartElement("v", NS);

			var d = c.dr.GetDecimal(ordinal);
			if (d.TryFormat(scratch.AsSpan(), out var sl))
			{
				w.WriteRaw(scratch, 0, sl);
			}

			w.WriteEndElement();
			w.WriteEndElement();
		}
	}

	sealed class DoubleFieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var w = c.xw;
			w.WriteStartElement("c", NS);

			w.WriteStartElement("v", NS);

			var d = c.dr.GetDouble(ordinal);
			if (d.TryFormat(scratch.AsSpan(), out var sl))
			{
				w.WriteRaw(scratch, 0, sl);
			}

			w.WriteEndElement();
			w.WriteEndElement();
		}
	}

	sealed class Int32FieldWriter : FieldWriter
	{
		public override void WriteField(Context c, int ordinal)
		{
			var w = c.xw;
			w.WriteStartElement("c", NS);

			w.WriteStartElement("v", NS);

			var val = c.dr.GetInt32(ordinal);
			if (val.TryFormat(scratch.AsSpan(), out var sl))
			{
				w.WriteRaw(scratch, 0, sl);
			}

			w.WriteEndElement();
			w.WriteEndElement();
		}
	}
}
