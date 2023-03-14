using System;

namespace Sylvan.Data.Excel;

static class HexCodec
{
	static readonly char[] HexMap = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'a', 'b', 'c', 'd', 'e', 'f' };

	internal static int ToHexCharArray(byte[] dataBuffer, int offset, int length, char[] outputBuffer, int outputOffset)
	{
		if (length * 2 > outputBuffer.Length - outputOffset)
			throw new ArgumentException();

		var idx = offset;
		var end = offset + length;
		for (; idx < end; idx++)
		{
			var b = dataBuffer[idx];
			var lo = HexMap[b & 0xf];
			var hi = HexMap[b >> 4];
			outputBuffer[outputOffset++] = hi;
			outputBuffer[outputOffset++] = lo;
		}
		return length * 2;
	}
}
