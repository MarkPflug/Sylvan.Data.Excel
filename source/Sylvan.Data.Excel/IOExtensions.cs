
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Sylvan
{
	static class IOExtensions
	{
		//public static Task<int> ReadAllAsync(this Stream stream, byte[] buffer, int offset, int length, CancellationToken cancel = default)
		//{
		//	return ReadAllAsync(stream, buffer.AsMemory().Slice(offset, length), cancel);
		//}

		//public static async Task<int> ReadAllAsync(this Stream stream, Memory<byte> mem, CancellationToken cancel = default)
		//{
		//	int read = 0;
		//	var len = mem.Length;
		//	while (read < len)
		//	{
		//		var c = await stream.ReadAsync(mem, cancel);
		//		if (c == 0) break;
		//		read += c;
		//		mem = mem.Slice(c);
		//	}
		//	return read;
		//}
	}
}
