#nullable enable
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;

namespace Sylvan.Data.Excel;

sealed class DefaultExcelSchema : ExcelSchemaProvider
{
	readonly bool hasHeaders;

	internal static DbColumn DefaultColumnSchema = new DefaultExcelSchemaColumn(string.Empty);

	sealed class DefaultExcelSchemaColumn : DbColumn
	{
		public DefaultExcelSchemaColumn(string? name)
		{
			this.ColumnName = name ?? string.Empty;
			this.DataType = typeof(string);
			this.DataTypeName = this.DataType.Name;
			this.AllowDBNull = true;
		}
	}

	public DefaultExcelSchema(bool hasHeaders)
	{
		this.hasHeaders = hasHeaders;
	}

	public override DbColumn? GetColumn(string sheetName, string? columName, int ordinal)
	{
		var name = hasHeaders ? columName : ExcelSchema.GetExcelColumnName(ordinal);
		return new DefaultExcelSchemaColumn(name);
	}

	public override bool HasHeaders(string sheetName)
	{
		return hasHeaders;
	}
}

/// <summary>
/// An implementation of IExcelSchemaProvider that allows defining per-column types.
/// </summary>
public sealed class ExcelSchema : ExcelSchemaProvider
{
	/// <summary>
	/// A schema that expects each sheet to have a header row, and describes
	/// each column as being a nullable string.
	/// </summary>
	public static IExcelSchemaProvider Default = new DefaultExcelSchema(true);

	/// <summary>
	/// A schema that does not expect each sheet to have a header row, and describes
	/// each column as being a nullable string. Column names are exposed as the Excel column header "A", "B", etc.
	/// </summary>
	public static IExcelSchemaProvider NoHeaders = new DefaultExcelSchema(false);

	class ExcelSchemaColumn : DbColumn
	{
		public ExcelSchemaColumn(string? name)
		{
			this.ColumnName = name ?? string.Empty;
		}
	}

	Dictionary<string, SheetInfo>? sheets;
	SheetInfo? defaultSchema;

	/// <summary>
	/// Creates a new ExcelSchema instance.
	/// </summary>
	/// <param name="hasHeaders">Indicates if the sheet contains a header row.</param>
	/// <param name="columns">The schema column definitions for the sheet.</param>
	public ExcelSchema(bool hasHeaders, IEnumerable<DbColumn> columns)
	{
		this.defaultSchema = new SheetInfo(string.Empty, hasHeaders, columns);
	}

	/// <summary>
	/// Creates a new ExcelSchema instance.
	/// </summary>
	public ExcelSchema()
	{
		this.defaultSchema = null;
	}

	class SheetInfo
	{
		public string Name { get; }
		public bool HasHeaders { get; }
		public DbColumn[] Columns { get; }

		public SheetInfo(string name, bool hasHeaders, IEnumerable<DbColumn> columns)
		{
			this.Name = name;
			this.HasHeaders = hasHeaders;
			this.Columns = columns.ToArray();
		}

		public DbColumn? GetColumn(string? name, int ordinal)
		{
			if (name != null)
			{
				foreach (var col in Columns)
				{
					var columnName = col.BaseColumnName ?? col.ColumnName;

					if (StringComparer.OrdinalIgnoreCase.Equals(columnName, name))
					{
						return col;
					}
				}
			}
			if (ordinal < Columns.Length)
			{
				var col = Columns[ordinal];

				if (col.BaseColumnName == null)
				{
					return Columns[ordinal];
				}
			}
			return null;
		}
	}

	/// <summary>
	/// Adds a schema for a specific sheet.
	/// </summary>
	/// <param name="sheetName">The name of the sheet the schema applies to.</param>
	/// <param name="hasHeaders">Incidates if the sheet has a header row.</param>
	/// <param name="columns">The schema column definitions for the sheet.</param>
	/// <exception cref="ArgumentNullException">If the sheet name is null.</exception>
	public ExcelSchema Add(string sheetName, bool hasHeaders, IEnumerable<DbColumn> columns)
	{
		if (sheetName == null) throw new ArgumentNullException(nameof(sheetName));
		if (sheets == null)
			sheets = new Dictionary<string, SheetInfo>(StringComparer.OrdinalIgnoreCase);

		var sheetInfo = new SheetInfo(sheetName, hasHeaders, columns);
		this.sheets.Add(sheetName, sheetInfo);
		return this;
	}

	/// <inheritdoc/>
	public override DbColumn? GetColumn(string sheetName, string? columnName, int ordinal)
	{
		if (sheets != null)
		{
			if (sheets.TryGetValue(sheetName, out SheetInfo? info))
			{
				var name = info.HasHeaders ? columnName : GetExcelColumnName(ordinal);
				var schema = info.GetColumn(name, ordinal);
				if (schema != null)
				{
					return schema;
				}
			}
		}

		if (defaultSchema != null)
		{
			var name = defaultSchema.HasHeaders ? columnName : GetExcelColumnName(ordinal);
			var schema = defaultSchema.GetColumn(name, ordinal);
			if (schema != null)
			{
				return schema;
			}
		}

		return DefaultExcelSchema.DefaultColumnSchema;
	}

	/// <inheritdoc/>
	public override bool HasHeaders(string sheetName)
	{
		if (sheets != null && sheets.TryGetValue(sheetName, out SheetInfo? info))
		{
			return info.HasHeaders;
		}
		if (defaultSchema != null)
		{
			return defaultSchema.HasHeaders;
		}
		return false;
	}

	internal static string GetExcelColumnName(int idx)
	{
		var col = idx + 1;
		string name = string.Empty;
		while (col > 0)
		{
			var i = col - 1;
			col = Math.DivRem(i, 26, out var rem);
			name = ((char)('A' + rem)) + name;
		}
		return name;
	}
}
