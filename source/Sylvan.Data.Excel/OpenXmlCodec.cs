using System.Text.RegularExpressions;

namespace Sylvan.Data.Excel;

// provides functions for en/decoding _xHHHH_ encoded characters
// This encoding is mentioned in ECMA-376 22.4.2.2 (bstr Basic String)
// TODO: this code could probably benefit from some optimization.
static class OpenXmlCodec
{	
	// replace most control characters as well as underscore characters.
	// the underscore replacement is to "escape" underscores that might otherwise be
	// seen as an encoded sequence.
	// This ends up encoding more than is technically required, as underscores only need to be encoded
	// if they would otherwise be detected as an escape sequence.
	static readonly Regex EncodeCharRegex = new Regex(@"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f_]", RegexOptions.Compiled);

	static readonly Regex DecodeCharRegex = new Regex(@"_x[0-9a-fA-F]{4}_", RegexOptions.Compiled);

	public static string EncodeString(string str)
	{
		return EncodeCharRegex.Replace(str, EncodeChar);
	}

	public static string DecodeString(string str)
	{
		return DecodeCharRegex.Replace(str, DecodeChar);
	}

	static readonly MatchEvaluator EncodeChar = EncodeReplace;
	static readonly MatchEvaluator DecodeChar = DecodeReplace;

	static string EncodeReplace(Match m)
	{
		int c = m.Value[0];
		var str =  $"_x{c:x4}_";
		return str;
	}

	static int GetHexValue(char c)
	{
		return
			c >= '0' && c <= '9'
			? c - '0'
			: c >= 'a' && c <= 'f'
			? c - 'a'
			: c >= 'A' && c <= 'F'
			? c - 'A'
			: throw new System.Exception();
	}

	static string DecodeReplace(Match m)
	{
		var str = m.Value;
		char c =
			(char) (
			(GetHexValue(str[2]) << 12) |
			(GetHexValue(str[3]) << 8) |
			(GetHexValue(str[4]) << 4) |
			(GetHexValue(str[5]) << 0)
			);
		return "" + c;
	}
}
