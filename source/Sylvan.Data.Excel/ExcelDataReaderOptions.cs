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
			this.GetNullAsEmptyString = true;
		}

		/// <summary>
		/// Gets or sets the schema for the data in the workbook.
		/// </summary>
		public IExcelSchemaProvider Schema { get; set; }

		/// <summary>
		/// Indicates if GetString will return an emtpy string
		/// when the underlying value is null. If false, GetString
		/// will throw an exception when the underlying value is null.
		/// The default is true.
		/// </summary>
		public bool GetNullAsEmptyString { get; set; }

		/// <summary>
		/// Indicates if GetString will throw an ExcelFormulaException or return a string value
		/// when accesing a cell containing a formula error.
		/// </summary>
		public bool GetErrorAsNull { get; set; }
	}
}
