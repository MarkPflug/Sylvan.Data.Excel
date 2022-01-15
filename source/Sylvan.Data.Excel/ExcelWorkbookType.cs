namespace Sylvan.Data.Excel
{
	/// <summary>
	/// The type of workbook.
	/// </summary>
	public enum ExcelWorkbookType
	{
		/// <summary>
		/// Represents an unknown workbook type.
		/// </summary>
		Unknown = 0,

		/// <summary>
		/// An .xls file.
		/// </summary>
		Excel = 1,

		/// <summary>
		/// An .xlsx file.
		/// </summary>
		ExcelXml = 2,

		/// <summary>
		/// An .xslb file.
		/// </summary>
		ExcelBinary = 3,
	}
}
