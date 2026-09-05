using Xunit;

namespace Sylvan.Data.Excel;

public class OpenXmlCodecTests
{
	[Theory]
	[InlineData("\x0003", "_x0003_")]
	[InlineData("a\x0003b", "a_x0003_b")]
	[InlineData("_x0003_", "_x005f_x0003_")]
	[InlineData("_X0003_", "_X0003_")]
	[InlineData("_", "_")]
	[InlineData("a\rb", "a_x000d_b")]
	[InlineData("a\r\nb", "a_x000d_\nb")]
	[InlineData("a\nb", "a\nb")]
	[InlineData("\r=1+2", "_x000d_=1+2")]
	[InlineData("_x000D_", "_x005f_x000D_")]
	[InlineData("_x005F_x0041_", "_x005f_x005F_x005f_x0041_")]
	public void Encode(string input, string expected)
	{
		var result = OpenXmlCodec.EncodeString(input);
		Assert.Equal(expected, result);
	}

	[Theory]
	[InlineData("_x0003_", "\x0003")]
	[InlineData("a_x0003_b", "a\x0003b")]
	[InlineData("_x005f_x0003_", "_x0003_")]
	[InlineData("_X0003_", "_X0003_")]
	[InlineData("_", "_")]
	[InlineData("a_x000d_b", "a\rb")]
	[InlineData("a_x000d_\nb", "a\r\nb")]
	[InlineData("a\nb", "a\nb")]
	[InlineData("_x000d_=1+2", "\r=1+2")]
	[InlineData("_x005f_x000D_", "_x000D_")]
	[InlineData("_x005f_x005F_x005f_x0041_", "_x005F_x0041_")]
	public void Decode(string input, string expected)
	{
		var result = OpenXmlCodec.DecodeString(input);
		Assert.Equal(expected, result);
	}
}
