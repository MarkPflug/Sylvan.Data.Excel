using System;
using System.Runtime.InteropServices;

namespace Sylvan.Data.Excel;

partial class ExcelDataReader
{
	internal enum FieldType : int
	{
		Null = 0,
		Numeric,
		DateTime,
		String,
		SharedString,
		Boolean,
		Error,
	}

	[StructLayout(LayoutKind.Explicit)]
	private protected struct FieldInfo
	{
		public static readonly FieldInfo Null = default;

		[FieldOffset(0)]
		internal string? strValue;

		[FieldOffset(8)]
		internal bool boolValue;
		[FieldOffset(8)]
		internal int ssIdx;
		[FieldOffset(8)]
		internal double numValue;
		[FieldOffset(8)]
		internal DateTime dtValue;

		[FieldOffset(16)]
		internal FieldType type;

		[FieldOffset(20)]
		internal int xfIdx;



		internal bool IsEmptyValue
		{
			get
			{
				return this.type == FieldType.Null || (this.type == FieldType.String && this.strValue?.Length == 0);
			}
		}

		internal ExcelErrorCode ErrorCode
		{
			get { return (ExcelErrorCode)numValue; }
		}

		internal bool BoolValue
		{
			get { return boolValue; }
		}

		public FieldInfo(string str)
		{
			this = default;
			this.type = FieldType.String;
			this.strValue = str;
		}

		public FieldInfo(bool b)
		{
			this = default;
			this.type = FieldType.Boolean;
			this.boolValue = b;
		}

		public FieldInfo(ExcelErrorCode c)
		{
			this = default;
			this.type = FieldType.Error;
			this.numValue = (double)c;
		}

		public FieldInfo(uint val, FieldType type)
		{
			this = default;
			this.numValue = val;
			this.type = type;
		}

		public FieldInfo(double val, ushort ifIdx)
		{
			this = default;
			this.type = FieldType.Numeric;
			this.numValue = val;
			this.xfIdx = ifIdx;
		}

#if DEBUG
		public override string ToString()
		{
			switch (type)
			{
				case FieldType.Numeric:
					return "Double: " + numValue;
				case FieldType.String:
					return "String: " + strValue;
				case FieldType.Boolean:
					return "Bool:" + boolValue;
			}
			return "NULL";
		}
#endif
	}
}
