namespace Sylvan.Data.Excel;

partial class ExcelDataReader
{
	private protected class SheetInfo
	{
		public SheetInfo(
			string name,
			bool hidden)
		{
			this.Name = this.Part = name;
			this.Hidden = hidden;
		}

		public SheetInfo(
			string name, 
			string part, 
			bool hidden)
		{
			this.Name = name;
			this.Part = part;
			this.Hidden = hidden;
		}

		public string Name { get; }
		public string Part { get; }
		public bool Hidden { get; }
	}
}
