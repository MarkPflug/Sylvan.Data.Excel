﻿using System.Data;
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
		// Test a reading values where a string is represented with an empty element.
		// this happens when formula calculation produce empty values
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

	[Fact]
	public void NoCountStyle()
	{
		// Test reading file produced by JasperReports, which doesn't write the count attribute
		// on xfCells

		var reader = XlsxBuilder.Create(TestData.EmptyString, null, TestData.NoCountStyle);
		// implicit assert that creating the reader doesn't throw.
		Assert.NotNull(reader);
	}

	[Fact]
	public void EmptyInlineStr()
	{
		// Test reading file produced by JasperReports, which writes inlineStr values
		// that are empty elements.

		var reader = XlsxBuilder.Create(TestData.InlineStringEmpty2);
		Assert.True(reader.Read());
		Assert.Equal(3, reader.RowFieldCount);
		Assert.Equal(string.Empty, reader.GetString(0));
		Assert.Equal(ExcelDataType.Null, reader.GetExcelDataType(0));
		Assert.Equal(string.Empty, reader.GetString(1));
		Assert.Equal(ExcelDataType.Null, reader.GetExcelDataType(1));
		Assert.Equal("c", reader.GetString(2));
		Assert.True(reader.Read());
		Assert.Equal("a", reader.GetString(0));
		Assert.Equal("b", reader.GetString(1));
		Assert.Equal("c", reader.GetString(2));
		Assert.False(reader.Read());
	}

	[Fact]
	public void EmptyTrailingRow()
	{
		// If the final (or trailing) row contains a shared string referencing
		// an empty string, treat it as a null/empty value.

		var reader = XlsxBuilder.Create(TestData.EmptySSTrailingRow, TestData.SharedStringEmpty);
		Assert.True(reader.Read());
		Assert.Equal(3, reader.RowFieldCount);
		Assert.Equal("a", reader.GetString(0));
		Assert.Equal("a", reader.GetString(1));
		Assert.False(reader.Read());
	}
}
