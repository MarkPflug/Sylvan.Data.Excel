using System.Text.RegularExpressions;

namespace Sylvan.Data.Excel;

static class OpenXmlCodec
{
	static readonly Regex EncodeCharRegex = new Regex(@"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", RegexOptions.Compiled);

	static readonly Regex DecodeCharRegex = new Regex(@"_x[0-9a-fA-F]{4}_", RegexOptions.Compiled);

	// TODO: need to encode certain underscore characters too.
	// when the input matches the DecodeCharRegex, the leading underscore needs to be
	// replaced.
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
