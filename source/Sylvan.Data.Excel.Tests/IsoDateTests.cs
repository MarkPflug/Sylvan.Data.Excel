using System;
using Xunit;

namespace Sylvan.Data.Excel;

public class IsoDateTests
{
	[Fact]
	public void Fact1()
	{
		var dt = new DateTime(2022, 6, 21, 13, 14, 15, DateTimeKind.Utc).AddTicks(1234567);
		var str = IsoDate.ToStringIso(dt);
		var l = str.Length;
		var str2 = IsoDate.ToDateStringIso(dt);
		Assert.Equal("2022-06-21T13:14:15.1234567Z", str);
	}
}
