namespace Sylvan.Data.Excel
{
	/// <summary>
	/// Represents in internal data types supported by Excel.
	/// </summary>
	public enum ExcelDataType
	{
		/// <summary>
		/// A cell that contains no value.
		/// </summary>
		Null = 0,
		/// <summary>
		/// A numeric value. This is also used to represent DateTime values.
		/// </summary>
		Numeric,
		/// <summary>
		/// A DateTime value. This is an uncommonly used representation in .xlsx files.
		/// </summary>
		DateTime,
		/// <summary>
		/// A text field.
		/// </summary>
		String,
		/// <summary>
		/// A formula cell that contains a boolean.
		/// </summary>
		Boolean,
		/// <summary>
		/// A formula cell that contains an error.
		/// </summary>
		Error,
	}
}
