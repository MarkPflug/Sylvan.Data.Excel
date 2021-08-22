using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;

namespace Sylvan.Data.Excel
{
	sealed class DefaultExcelSchema : IExcelSchemaProvider
	{
		readonly bool hasHeaders;

		internal static DbColumn DefaultColumnSchema = new DefaultExcelSchemaColumn("");

		sealed class DefaultExcelSchemaColumn : DbColumn
		{
			public DefaultExcelSchemaColumn(string? name)
			{
				this.ColumnName = name ?? "";
				this.DataType = typeof(string);
				this.DataTypeName = "string";
				this.AllowDBNull = true;
			}
		}

		public DefaultExcelSchema(bool hasHeaders)
		{
			this.hasHeaders = hasHeaders;
		}

		public DbColumn? GetColumn(string sheetName, string? columName, int ordinal)
		{
			var name = hasHeaders ? columName : ExcelSchema.GetExcelColumnName(ordinal);
			return new DefaultExcelSchemaColumn(name);

		}

		public bool HasHeaders(string sheetName)
		{
			return hasHeaders;
		}
	}

	public class ExcelSchema : IExcelSchemaProvider
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

		public ExcelSchema(bool hasHeaders, IEnumerable<DbColumn> columns)
		{
			this.defaultSchema = new SheetInfo(string.Empty, hasHeaders, columns);
		}

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
				throw new NotImplementedException();
			}
		}

		public ExcelSchema Add(string sheetName, bool hasHeaders, IEnumerable<DbColumn> columns)
		{
			if (sheetName == null) throw new ArgumentNullException(nameof(sheetName));
			if (sheets == null)
				sheets = new Dictionary<string, SheetInfo>(StringComparer.OrdinalIgnoreCase);

			var sheetInfo = new SheetInfo(sheetName, hasHeaders, columns);
			this.sheets.Add(sheetName, sheetInfo);
			return this;
		}

		public DbColumn? GetColumn(string sheetName, string? columnName, int ordinal)
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

		public bool HasHeaders(string sheetName)
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
			string name = "";
			while (col > 0)
			{
				var i = col - 1;
				col = Math.DivRem(i, 26, out var rem);
				name = ((char)('A' + rem)) + name;
			}
			return name;
		}
	}
}
