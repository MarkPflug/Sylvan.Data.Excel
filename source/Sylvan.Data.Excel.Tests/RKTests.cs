using System;
using Xunit;

namespace Sylvan.Data.Excel;

public class RKTests
{
	static int ToRk(int val)
	{
		return val << 2 | 0x2;
	}

	[Fact]
	public  void Test1()
	{
		unchecked
		{
			if (!Test(0x1fffffff))
			{

			}
			var v = (int)0xefffffff;
			if (!Test(v))
			{

			}
			v = (int)0x9fffffff;
			if (!Test(v))
			{

			}

			for (int i = 0; i < 32; i++)
			{
				var x = 1 << i;
				if (!Test(x))
				{

				}
				if (!Test(-x + 1))
				{

				}
			}
		}
	}

	bool Test(int val)
	{
		var rk = ToRk(val);
		var result = ExcelDataReader.GetRKVal(rk);
		return val == result;
	}
}
