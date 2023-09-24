using System;

namespace Sylvan.Data.Excel;

partial class ExcelDataReader
{
	private protected struct FieldInfo
	{
		public static readonly FieldInfo Null = default;

		public ExcelDataType type;
		public bool isSS;
		public string? strValue;
		public int ssIdx;
		public double numValue;
		public DateTime dtValue;
		public int xfIdx;

		internal ExcelErrorCode ErrorCode
		{
			get { return (ExcelErrorCode)numValue; }
		}

		internal bool BoolValue
		{
			get { return numValue != 0d; }
		}

		public FieldInfo(string str)
		{
			this = default;
			this.type = ExcelDataType.String;
			this.strValue = str;
		}

		public FieldInfo(bool b)
		{
			this = default;
			this.type = ExcelDataType.Boolean;
			this.numValue = b ? 1 : 0;
		}

		public FieldInfo(ExcelErrorCode c)
		{
			this = default;
			this.type = ExcelDataType.Error;
			this.numValue = (double)c;
		}

		public FieldInfo(uint val, ExcelDataType type)
		{
			this = default;
			this.numValue = val;
			this.type = type;
		}

		public FieldInfo(double val, ushort ifIdx)
		{
			this = default;
			this.type = ExcelDataType.Numeric;
			this.numValue = val;
			this.xfIdx = ifIdx;
		}

		public override string ToString()
		{
			switch (type)
			{
				case ExcelDataType.Numeric:
					return "Double: " + numValue;
				case ExcelDataType.String:
					return "String: " + strValue;
			}
			return "NULL";
		}
	}
}
