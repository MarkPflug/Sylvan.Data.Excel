#nullable enable
using System.Data.Common;

namespace Sylvan.Data.Excel;

/// <summary>
/// A base implementation of IExcelSchemaProvider
/// </summary>
public abstract class ExcelSchemaProvider : IExcelSchemaProvider
{
	/// <inheritdoc/>
	public abstract DbColumn? GetColumn(string sheetName, string? name, int ordinal);

	/// <inheritdoc/>
	public abstract bool HasHeaders(string sheetName);

	/// <inheritdoc/>
	public virtual int GetFieldCount(ExcelDataReader reader)
	{
		return reader.RowFieldCount;
	}	
}

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
	/// Called to get the number of fields in the worksheet.
	/// </summary>
	/// <param name="reader">The reader being initialized.</param>
	/// <returns>The number of fields.</returns>
	int GetFieldCount(ExcelDataReader reader)
#if NETSTANDARD2_1_OR_GREATER
	{
		return reader.RowFieldCount;
	}
#else
	; // abstract
#endif

	/// <summary>
	/// Called to determine the schema for a column in a worksheet.
	/// </summary>
	/// <param name="sheetName">The name of the worksheet.</param>
	/// <param name="name">The name of the column</param>
	/// <param name="ordinal">The ordinal position of the column.</param>
	/// <returns>A DbColumn that defines the schema for the column.</returns>
	DbColumn? GetColumn(string sheetName, string? name, int ordinal);
}
