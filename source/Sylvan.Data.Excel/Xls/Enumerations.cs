namespace Sylvan.Data.Excel;

enum RecordType
{
	Dimension = 0x0200,
	YearEpoch = 0x022,
	Blank = 0x0201,
	Number = 0x0203,
	Label = 0x0204,
	BoolErr = 0x0205,
	Formula = 0x0006,
	String = 0x0207,

	BOF = 0x0809,

	Continue = 0x003c,
	CRN = 0x005a,
	LabelSST = 0x00fd,

	RK = 0x027e,

	MulRK = 0x00BD,
	EOF = 0x000A,
	XF = 0x00e0,

	Font = 0x0031,
	ExtSst = 0x00ff,
	Format = 0x041e,
	Style = 0x0293,
	Row = 0x0208,

	ExternSheet = 0x0017,
	DefinedName = 0x0018,
	Country = 0x008c,

	Index = 0x020B,

	CalcCount = 0x000c,
	CalcMode = 0x000d,
	Precision = 0x000e,
	RefMode = 0x000f,

	Delta = 0x0010,
	Iteration = 0x0011,
	Protect = 0x0012,
	Password = 0x0013,
	Header = 0x0014,
	Footer = 0x0015,
	ExternCount = 0x0016,

	Guts = 0x0080,
	SheetPr = 0x0081,
	GridSet = 0x0082,
	HCenter = 0x0083,
	VCenter = 0x0084,
	Sheet = 0x0085,
	WriteProt = 0x0086,

	Sort = 0x0090,

	ColInfo = 0x007d,

	Sst = 0x00fc,
	MulBlank = 0x00be,
	RString = 0x00d6,
	Array = 0x0221,
	SharedFmla = 0x04bc,
	DataTable = 0x0236,
	DBCell = 0x00d7,
}

enum BOFType
{
	WorkbookGlobals = 0x0005,
	VisualBasicModule = 0x0006,
	Worksheet = 0x0010,
	Chart = 0x0020,
	Biff4MacroSheet = 0x0040,
	Biff4WorkbookGlobals = 0x0100,
}