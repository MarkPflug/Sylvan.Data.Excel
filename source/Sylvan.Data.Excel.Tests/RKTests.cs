using Xunit;

namespace Sylvan.Data.Excel;

public class RKTests
{

	[Fact]
	public void Test1()
	{
		unchecked
		{			
			//var d1 = ExcelDataReader.GetRKVal((int)0xc1e00000);
			
			//var d2 = ExcelDataReader.GetRKVal((int)0x40000001);
			//var d3 = ExcelDataReader.GetRKVal((int)0x29a023ca);
			var d4 = ExcelDataReader.GetRKVal((int)0xd65fdc3a);

		}
	}
}
