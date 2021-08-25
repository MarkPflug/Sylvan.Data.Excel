using System;
using System.Collections.Generic;

namespace Sylvan.Data.Excel
{
	public enum FormatKind
	{
		String = 1,
		Number,
		Date,
		Time,
	}

	public enum ExcelErrorCode
	{
		Null = 0,
		DivideByZero = 7,
		Value = 15,
		Reference = 23,
		Name = 29,
		Number = 36,
		NotAvailable = 42,
	}

	public sealed class ExcelFormulaException : Exception
	{
		internal ExcelFormulaException(int col, int row, ExcelErrorCode code)
		{
			this.Row = row;
			this.Column = col;
			this.ErrorCode = code;
		}

		public int Row { get; }
		public int Column { get; }
		public ExcelErrorCode ErrorCode { get; }
	}

	public sealed class ExcelFormat
	{
		internal static Dictionary<int, ExcelFormat> CreateFormatCollection()
		{
			var dict = new Dictionary<int, ExcelFormat>();
			for (int i = 0; i < standardFormats.Length; i++)
			{
				var fmt = standardFormats[i];
				if (fmt != null)
					dict.Add(i, fmt);
			}
			return dict;
		}

		static FormatKind DetermineKind(string spec)
		{
			// TODO: this whole function could use some cleanup/rework.
			// passes test cases for now at least.

			bool hasTimeElements = false;
			bool hasNumberElements = false;

			int count;
			for (int i = 0; i < spec.Length; i++)
			{
				var c = spec[i];
				c = char.ToLowerInvariant(c);
				switch (c)
				{
					case '[':
						for (var j = i + 1; j < spec.Length; j++)
						{
							c = spec[j];
							if (c == ']')
							{
								i = j;
								break;
							}
						}
						break;
					case 'a':
					case 'p':
						if (i + 1 < spec.Length)
						{
							c = char.ToLowerInvariant(spec[i + 1]);
							if (c == 'm')
							{
								i++;
							}
						}
						break;
					case '\\':
						i++;
						break;
					case '0':
					case '#':
						hasNumberElements = true;
						break;
					case 'y':
					case 'd':
						return FormatKind.Date;
					case ':':
						i++;
						count = 0;
						for (; i < spec.Length; i++)
						{
							c = spec[i];
							if(c == 'm')
							{
								count++;
								continue;
							}
						}
						if(count > 0)
						{
							hasTimeElements = true;
						}
						break;
					case 'm':
						count = 1;
						bool time = false;
						for (var j = i + 1; j < spec.Length; j++)
						{
							c = spec[j];
							if (c == 'm')
							{
								count++;
								i = j;
								continue;
							}
							if (c == ':')
							{
								i = j;
								hasTimeElements = true;
								time = true;
								break;
							}
							break;
						}
						if (time)
							break;

						return FormatKind.Date;
					case 'h':
					case 's':
						hasTimeElements = true;
						break;
				}
			}
			if (hasTimeElements)
				return FormatKind.Time;
			if (hasNumberElements)
				return FormatKind.Number;
			return FormatKind.String;
		}

		internal ExcelFormat(string spec)
		{
			this.Format = spec;
			this.Kind = DetermineKind(spec);
			//this.format = FormatKind switch
			//{
			//	FormatKind.Date => "o",
			//	FormatKind.Time => "HH:mm:ss.FFFFFFF",
			//	_ => "G",
			//};
		}

		internal ExcelFormat(string spec, FormatKind kind, string? format = null)
		{
			this.Format = spec;
			this.Kind = kind;
		}

		/// <summary>
		/// Gets the format string.
		/// </summary>
		public string Format { get; private set; }

		/// <summary>
		/// Gets the kind of value the format string specifies.
		/// </summary>
		public FormatKind Kind { get; private set; }

		internal string FormatValue(double value, int dateOffset = 1900)
		{
			var kind = this.Kind;
			switch (kind)
			{
				case FormatKind.Number:
					return value.ToString("G");
				case FormatKind.Date:
				case FormatKind.Time:
					if (ExcelDataReader.TryGetDate(value, dateOffset, out var dt))
					{
						var fmt =
						dt.TimeOfDay == TimeSpan.Zero
						? "yyyy-MM-dd" // omit rendering the time when the value is midnight
						: "yyyy-MM-ddTHH:mm:ss.FFFFFFF";

						return dt.ToString(fmt);
					}
					else
					{
						// for values rendered as time (not including date) that are in the
						// range 0-1 (which renders in Excel as 1900-01-00),
						// allow these to be reported as just the time component.
						if (value < 1d && value >= 0d && Kind == FormatKind.Time)
						{
							// omit rendering the date when the value is in the range 0-1
							// this would render in Excel as the date 
							var fmt = "HH:mm:ss.FFFFFF";
							dt = DateTime.MinValue.AddDays(value);
							return dt.ToString(fmt);
						}
					}
					// We arrive here for negative values which render in Excel as "########" (not meaningful)
					// or 1900-01-00 date, which isn't a valid date.
					// or 1900-02-29, which is a non-existent date.
					// The value can still be accessed via GetDouble.
					return string.Empty;
			}
			return value.ToString();
		}

		static readonly ExcelFormat[] standardFormats;

		static ExcelFormat()
		{
			var fmts = new ExcelFormat[50];
			fmts[0] = new ExcelFormat("General", FormatKind.Number, "G");
			fmts[1] = new ExcelFormat("0", FormatKind.Number);
			fmts[2] = new ExcelFormat("0.00", FormatKind.Number);
			fmts[3] = new ExcelFormat("#,##0", FormatKind.Number);
			fmts[4] = new ExcelFormat("#,##0.00", FormatKind.Number);
			// 5: "($#: //##0_);($#: //##0)"
			// 6: "($#: //##0_);[Red]($#: //##0)"
			// 7: "($#: //##0.00_);($#: //##0.00)"
			// 8: "($#: //##0.00_);[Red]($#: //##0.00)"
			fmts[9] = new ExcelFormat("0%", FormatKind.Number);
			fmts[10] = new ExcelFormat("0.00%", FormatKind.Number);
			fmts[11] = new ExcelFormat("0.00E+00", FormatKind.Number, "0.00E+0");
			// 12: "# ?/?"
			// 13: "# ??/??"
			fmts[14] = new ExcelFormat("m/d/yy", FormatKind.Date);
			fmts[15] = new ExcelFormat("d-mmm-yy", FormatKind.Date);
			fmts[16] = new ExcelFormat("d-mmm", FormatKind.Date);
			fmts[17] = new ExcelFormat("mmm-yy", FormatKind.Date);
			fmts[18] = new ExcelFormat("h:mm AM/PM", FormatKind.Time, "h:mm tt");
			fmts[19] = new ExcelFormat("h:mm:ss AM/PM", FormatKind.Time, "h:mm:ss tt");
			fmts[20] = new ExcelFormat("h:mm", FormatKind.Time, "h:mm");
			fmts[21] = new ExcelFormat("h:mm:ss", FormatKind.Time, "h:mm:ss");
			fmts[22] = new ExcelFormat("m/d/yy h:mm:ss", FormatKind.Date, "m/d/yy h:mm:ss");

			fmts[37] = new ExcelFormat("#,##0 ;(#,##0)", FormatKind.Number, "#,##0;(#,##0)");
			fmts[38] = new ExcelFormat("#,##0 ;[Red](#,##0)", FormatKind.Number, "#,##0;(#,##0)");
			fmts[39] = new ExcelFormat("#,##0.00;(#,##0.00)", FormatKind.Number, "#,##0.00;(#,##0.00)");
			fmts[40] = new ExcelFormat("#,##0.00;[Red](#,##0.00)", FormatKind.Number, "#,##0.00;(#,##0.00)");

			//	41: "_(* #: //##0_);_(* #: //##0);(* \" - \"_);_(@_)"
			//	42: "_($* #: //##0_);_($* #: //##0);($* \" - \"_);_(@_)"
			//	43: "_(* #: //##0.00_);_(* #: //##0.00);(* \" - \"_);_(@_)"
			//	44: "_($* #: //##0.00_);_($* #: //##0.00);($* \" - \"_);_(@_)"
			fmts[45] = new ExcelFormat("mm:ss", FormatKind.Time, "mm:ss");
			fmts[46] = new ExcelFormat("[h]:mm:ss", FormatKind.Time, "h:mm:ss");
			//	47: "mm:ss.0"
			fmts[48] = new ExcelFormat("##0.0E+0", FormatKind.Number);
			fmts[49] = new ExcelFormat("@", FormatKind.String);
			standardFormats = fmts;
		}

		public override string ToString()
		{
			return this.Format + " (" + this.Kind + ")";
		}
	}
}
