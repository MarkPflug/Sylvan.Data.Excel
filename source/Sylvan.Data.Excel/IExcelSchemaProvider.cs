#nullable enable
using System.Data.Common;

namespace Sylvan.Data.Excel;

/// <summary>
/// Provides schema information for an Excel data file.
/// </summary>
public interface IExcelSchemaProvider
{
	/// <summary>
	/// Called to determine if a worksheet contains a header row.
	/// </summary>
	/// <param name="sheetName">The name of the worksheet.</param>
	/// <returns>True if the first row should be interpreted as column headers.</returns>
	bool HasHeaders(string sheetName);

	/// <summary>
	/// Called to determine the schema for a column in a worksheet.
	/// </summary>
	/// <param name="sheetName">The name of the worksheet.</param>
	/// <param name="name"></param>
	/// <param name="ordinal"></param>
	/// <returns></returns>
	DbColumn? GetColumn(string sheetName, string? name, int ordinal);
}
