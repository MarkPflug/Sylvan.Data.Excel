﻿using System;
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
	public void NoCountSharedStrings()
	{
		// Test reading file produced by Crystal Reports, which doesn't write the
		// count/unique attribute on the shared strings table

		var reader = XlsxBuilder.Create(TestData.NoCountSSTWS, TestData.NoCountSST);
		Assert.True(reader.Read());
		Assert.Equal("a", reader.GetString(0));
		Assert.Equal("b", reader.GetString(1));
		Assert.False(reader.Read());
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
		// In the case that there is a trailing row that references empty strings
		// we will read it as a valid row. This should be uncommon enough that it won't affect anyone
		// in practice.

		var reader = XlsxBuilder.Create(TestData.EmptySSTrailingRow, TestData.SharedStringEmpty);
		Assert.True(reader.Read());
		Assert.Equal(3, reader.RowFieldCount);
		Assert.Equal("a", reader.GetString(0));
		Assert.Equal("a", reader.GetString(1));

		Assert.True(reader.Read());

		for (int i = 0; i < reader.RowFieldCount; i++)
		{
			Assert.True(reader.IsDBNull(i));
			Assert.Equal("", reader.GetString(i));

		}

		Assert.False(reader.Read());
	}

	[Fact]
	public void DateFormat()
	{
		var edr = XlsxBuilder.Create(
			TestData.DateFormatData, 
			null, 
			TestData.DateFormatStyle, 
			o => o.DateTimeFormat = "dd/MM/yyyy"
			);

		Assert.True(edr.Read());

		Assert.Equal(new DateTime(2022, 6, 9), edr.GetDateTime(0));
		Assert.Equal(ExcelDataType.Numeric, edr.GetExcelDataType(0));
		Assert.Equal("2022-06-09", edr.GetValue(0));

		Assert.Equal(new DateTime(2022, 6, 9), edr.GetDateTime(1));
		Assert.Equal(ExcelDataType.String, edr.GetExcelDataType(1));
		Assert.Equal("09/06/2022", edr.GetValue(1));

	}

	[Fact]
	public void DateFormat2()
	{
		var schema = Schema.Parse("a:date,b:date{dd/MM/yyyy}");
		var edr = XlsxBuilder.Create(
			TestData.DateFormatData,
			null,
			TestData.DateFormatStyle,
			o => o.Schema = new ExcelSchema(false, schema)
			);

		Assert.True(edr.Read());

		var value = new DateTime(2022, 6, 9);

		Assert.Equal(value, edr.GetDateTime(0));
		Assert.Equal(ExcelDataType.Numeric, edr.GetExcelDataType(0));
		// applying a schema returns a DateTime value
		Assert.Equal(value, edr.GetValue(0));

		Assert.Equal(new DateTime(2022, 6, 9), edr.GetDateTime(1));
		Assert.Equal(ExcelDataType.String, edr.GetExcelDataType(1));
		Assert.Equal(value, edr.GetValue(1));
	}

	[Fact]
	public void WhitespaceSharedStrings()
	{
		var r = XlsxBuilder.Create(TestData.WSSheet, TestData.WSSharedString, null, o => o.Schema = ExcelSchema.NoHeaders);

		Assert.True(r.Read());
		Assert.Equal(" a ", r.GetString(0));
		Assert.Equal("b", r.GetString(1));
		Assert.Equal(" a ", r.GetString(2));
		Assert.Equal("b", r.GetString(3));
		Assert.Equal(" a ", r.GetString(4));
		Assert.Equal("b", r.GetString(5));
	}

	[Fact]
	public void MalformedCellRef()
	{
		using var r = XlsxBuilder.Create(TestData.MalformedRef, null, null, o => o.Schema = ExcelSchema.NoHeaders);
		r.Read();
		Assert.Equal(0, r.GetInt32(0));
		// the cell reference "/1" is broken
		// and will be interpreted as the "next" cell instead.
		Assert.Equal(1, r.GetInt32(1));
	}

	[Fact]
	public void Phonetic()
	{
		using var r = XlsxBuilder.Create(TestData.Phonetic, TestData.PhoneticSST, null, o => o.Schema = ExcelSchema.NoHeaders);
		r.Read();
		Assert.Equal("NAME", r.GetString(0));
		Assert.Equal("A", r.GetString(1));
		Assert.Equal("C,D", r.GetString(2));
	}
}
