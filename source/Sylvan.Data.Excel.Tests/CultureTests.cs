using System;
using System.Globalization;
using Xunit;

namespace Sylvan.Data.Excel;

partial class XlsxTests
{
	[Fact]
	[UseCulture("it-it")]
	public void NumberInvariant()
	{
		var f = GetFile("Numbers");

		var opt = 
			new ExcelDataReaderOptions { 
				Schema = ExcelSchema.NoHeaders
				// culture defaults to invariant
			};
		var edr = ExcelDataReader.Create(f, opt);

		Assert.True(edr.Read());

		var v = edr.GetDouble(0);
		var s = edr.GetString(0);
		var o = edr.GetValue(0);

		Assert.Equal(3.3, v);
		// even though the thread culture (italian) uses  a comma for a decimal separator
		// the EDR returns the invariant as that is the default culture used by the options
		Assert.Equal("3.3", s);
		Assert.Equal("3.3", o);
	}

	[Fact]
	public void NumberCulture()
	{
		var f = GetFile("Numbers");

		var c = CultureInfo.GetCultureInfoByIetfLanguageTag("it-it");
		var opt = new ExcelDataReaderOptions { Schema = ExcelSchema.NoHeaders, Culture = c };
		var edr = ExcelDataReader.Create(f, opt);

		Assert.True(edr.Read());

		var v = edr.GetDouble(0);
		var s = edr.GetString(0);
		var o = edr.GetValue(0);

		Assert.Equal(3.3, v);
		// providing the italian culture causes string values to be formatted using
		// comma separator
		Assert.Equal("3,3", s);
		Assert.Equal("3,3", o);
	}


	[Fact]
	[UseCulture("it-it")]
	public void DateTimeInvariant()
	{
		var f = GetFile("DateTime");

			// culture defaults to invariant
		var opt = new ExcelDataReaderOptions();
		var edr = ExcelDataReader.Create(f, opt);

		for (int i = 0; i < 12; i++)
		{
			Assert.True(edr.Read());
		}

		var v = edr.GetDateTime(2);
		var s = edr.GetString(2);
		var o = edr.GetValue(2);

		Assert.Equal(new DateTime(1900, 1, 1, 2, 24, 0, DateTimeKind.Unspecified ), v);
		//// even though the thread culture (italian) uses  a comma for a decimal separator
		//// the EDR returns the invariant as that is the default culture used by the options
		Assert.Equal("1900-01-01T02:24:00", s);
		Assert.Equal("1900-01-01T02:24:00", o);
	}

	[Fact]
	public void DateTimeCulture()
	{
		var f = GetFile("DateTime");

		var c = CultureInfo.GetCultureInfoByIetfLanguageTag("it-it");
		var opt = new ExcelDataReaderOptions { Culture = c };
		var edr = ExcelDataReader.Create(f, opt);

		for (int i = 0; i < 12; i++)
		{
			Assert.True(edr.Read());
		}

		var v = edr.GetDateTime(2);
		var s = edr.GetString(2);
		var o = edr.GetValue(2);

		Assert.Equal(new DateTime(1900, 1, 1, 2, 24, 0, DateTimeKind.Unspecified), v);
		//// even though the thread culture (italian) uses  a comma for a decimal separator
		//// the EDR returns the invariant as that is the default culture used by the options
		Assert.Equal("01/01/1900 02:24:00", s);
		Assert.Equal("01/01/1900 02:24:00", o);
	}
}
