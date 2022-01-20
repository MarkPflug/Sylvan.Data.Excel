namespace Sylvan.Data.Excel
{
	/// <summary>
	/// Options for controlling the behavior of an <see cref="ExcelDataReader"/>.
	/// </summary>
	public sealed class ExcelDataReaderOptions
	{
		internal static readonly ExcelDataReaderOptions Default = new ExcelDataReaderOptions();

		/// <summary>
		/// Creates a new ExcelDataReaderOptions with the default values.
		/// </summary>
		public ExcelDataReaderOptions()
		{
			this.Schema = ExcelSchema.Default;
		}

		/// <summary>
		/// Gets or sets the schema for the data in the workbook.
		/// </summary>
		public IExcelSchemaProvider Schema { get; set; }

		/// <summary>
		/// Indicates if a cell will appear null or throw an ExcelFormulaException when accesing a cell containing a formula error.
		/// </summary>
		public bool GetErrorAsNull { get; set; }
	}
}
