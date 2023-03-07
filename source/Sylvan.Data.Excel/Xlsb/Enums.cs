namespace Sylvan.Data.Excel.Xlsb;


enum CellType
{
	Numeric,
	String,
	SharedString,
	Boolean,
	Error,
	Date,
}

enum RecordType
{
	Row = 0,
	CellBlank = 1,
	CellRK = 2,
	CellError = 3,
	CellBool = 4,
	CellReal = 5,
	CellSt = 6,
	CellIsst = 7,
	CellFmlaString = 8,
	CellFmlaNum = 9,
	CellFmlaBool = 10,
	CellFmlaError = 11,
	SSTItem = 19,
	Fmt = 44,
	XF = 47,
	BundleBegin = 143,
	BundleEnd = 144,
	BundleSheet = 156,
	BookBegin = 131,
	BookEnd = 132,
	Dimension = 148,
	SSTBegin = 159,
	SSTEnd = 160,
	StyleBegin = 278,
	StyleEnd = 279,
	CellXFStart = 617,
	CellXFEnd = 618,
	FmtStart = 615,
	FmtEnd = 616,
	SheetStart = 129,
	SheetEnd = 130,
	DataStart = 145,
	DataEnd = 146,
}