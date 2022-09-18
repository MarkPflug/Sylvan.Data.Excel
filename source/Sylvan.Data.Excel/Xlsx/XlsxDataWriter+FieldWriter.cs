using System;
using System.Data.Common;
using System.Xml;

namespace Sylvan.Data.Excel.Xlsx;

partial class XlsxDataWriter
{
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

		public void WriteField(XlsxDataWriter dw, XmlWriter w, DbDataReader r, int ordinal)
		{
			w.WriteStartElement("c", NS);
			Element(w);

			w.WriteStartElement("v", NS);
			WriteValue(dw, w, r, ordinal);
			w.WriteEndElement();
			w.WriteEndElement();
		}

		protected virtual void Element(XmlWriter w)
		{

		}

		internal abstract void WriteValue(XlsxDataWriter dw, XmlWriter w, DbDataReader dr, int idx);
	}

	sealed class StringFieldWriter : FieldWriter
	{


		protected override void Element(XmlWriter w)
		{
			w.WriteStartAttribute("t");
			w.WriteValue("s");
			w.WriteEndAttribute();
		}

		internal override void WriteValue(XlsxDataWriter dw, XmlWriter w, DbDataReader dr, int idx)
		{
			var s = dr.GetString(idx);
			var ssIdx = dw.sharedStrings.GetString(s);
			w.WriteValue(ssIdx);
		}
	}

	sealed class DateTimeFieldWriter : FieldWriter
	{
		
		protected override void Element(XmlWriter w)
		{
			w.WriteStartAttribute("t");
			w.WriteValue("d");
			w.WriteEndAttribute();

			w.WriteStartAttribute("s");

			var fmtId = "2";
			w.WriteValue(fmtId);
			w.WriteEndAttribute();
		}

		internal override void WriteValue(XlsxDataWriter dw, XmlWriter w, DbDataReader dr, int idx)
		{
			var dt = dr.GetDateTime(idx);
			if (IsoDate.TryFormatIso(dt, scratch.AsSpan(), out var sl))
			{
				w.WriteRaw(scratch, 0, sl);
			}
		}
	}

	sealed class DecimalFieldWriter : FieldWriter
	{
		
		internal override void WriteValue(XlsxDataWriter dw, XmlWriter w, DbDataReader dr, int idx)
		{
			var d = dr.GetDecimal(idx);
			if (d.TryFormat(scratch.AsSpan(), out var sl))
			{
				w.WriteRaw(scratch, 0, sl);
			}
		}
	}

	sealed class DoubleFieldWriter : FieldWriter
	{

		internal override void WriteValue(XlsxDataWriter dw, XmlWriter w, DbDataReader dr, int idx)
		{
			var d = dr.GetDouble(idx);
			if (d.TryFormat(scratch.AsSpan(), out var sl))
			{
				w.WriteRaw(scratch, 0, sl);
			}
		}
	}

	sealed class Int32FieldWriter : FieldWriter
	{

		internal override void WriteValue(XlsxDataWriter dw, XmlWriter w, DbDataReader dr, int idx)
		{
			var val = dr.GetInt32(idx);
			if (val.TryFormat(scratch.AsSpan(), out var sl))
			{
				w.WriteRaw(scratch, 0, sl);
			}
		}
	}

}
