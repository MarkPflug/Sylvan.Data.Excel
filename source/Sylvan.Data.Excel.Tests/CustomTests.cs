using System.Data;
using Xunit;

namespace Sylvan.Data.Excel;

public class CustomTests
{
	[Fact]
	public void InlineString()
	{
		// This tests the behavior of cells that use inlineString.
		// Excel doesn't create these files, as it will always put strings in the shared strings table.
		// I believe Apache POI (and maybe NPOI) will create such files though.
		var reader = XlsxBuilder.Create(TestData.InlineString);

		Assert.True(reader.Read());
		Assert.Equal("a", reader.GetString(0));
		Assert.Equal("b", reader.GetString(1));
		Assert.Equal("c", reader.GetString(2));
		Assert.False(reader.Read());
	}

	[Fact]
	public void InlineStringEmpty()
	{
		// Tests other odd cases with inlineString
		var reader = XlsxBuilder.Create(TestData.InlineStringEmpty);

		Assert.True(reader.Read());
		Assert.Equal("a", reader.GetString(0));
		Assert.Equal("", reader.GetString(1));
		Assert.Equal("", reader.GetString(2));
		Assert.Equal("d", reader.GetString(3));
		Assert.False(reader.Read());
	}

	[Fact]
	public void SharedStringRich()
	{
		// Tests formatted shared string text
		var reader = XlsxBuilder.Create(TestData.WorksheetRich, TestData.SharedStringRich);

		Assert.True(reader.Read());
		Assert.Equal("a", reader.GetString(0));
		Assert.Equal("abc", reader.GetString(1));
		Assert.Equal("c", reader.GetString(2));
		Assert.False(reader.Read());
	}

	[Fact]
	public void MissingStyle()
	{
		// Tests formatted shared string text
		var reader = XlsxBuilder.Create(TestData.UnknownFormat);

		Assert.True(reader.Read());
		Assert.Equal("1", reader.GetString(0));
		Assert.Equal("2", reader.GetString(1));
		Assert.Equal("3", reader.GetString(2));
		Assert.False(reader.Read());
	}

	[Fact]
	public void EmptySharedStringValue()
	{
		// Test a degenerate case produced by AG grid export to excel.
		var reader = XlsxBuilder.Create(TestData.EmptyValue, TestData.SharedStringSimple);
		Assert.True(reader.Read());
		// <v>0</v>
		Assert.Equal("a", reader.GetString(0));
		//<v></v> 
		Assert.True(reader.IsDBNull(1));
		Assert.Equal(string.Empty, reader.GetString(1));
		//<v/> 
		Assert.True(reader.IsDBNull(2));
		Assert.Equal(string.Empty, reader.GetString(2));
		Assert.False(reader.Read());
	}

	[Fact]
	public void EmptyString()
	{
		// Test a degenerate case produced by AG grid export to excel.
		var reader = XlsxBuilder.Create(TestData.EmptyString);
		Assert.True(reader.Read());
		Assert.False(reader.IsDBNull(0));
		Assert.Equal("a", reader.GetString(0));
		// <v/>
		Assert.True(reader.IsDBNull(1));
		Assert.Equal("", reader.GetString(1));
		// <v></v> 
		Assert.True(reader.IsDBNull(2));
		Assert.Equal("", reader.GetString(2));
		Assert.False(reader.Read());
	}
}
