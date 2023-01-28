using Xunit;

namespace Sylvan.Data.Excel;

public class OpenXmlCodecTests
{
	[Theory]
	[InlineData("\x0003", "_x0003_")]
	[InlineData("a\x0003b", "a_x0003_b")]
	public void Encode(string input, string expected)
	{
		var result = OpenXmlCodec.EncodeString(input);
		Assert.Equal(expected, result);

	}

	[Theory]
	[InlineData("_x0003_", "\x0003")]
	[InlineData("a_x0003_b", "a\x0003b")]
	public void Decode(string input, string expected)
	{
		var result = OpenXmlCodec.DecodeString(input);
		Assert.Equal(expected, result);

	}
}
