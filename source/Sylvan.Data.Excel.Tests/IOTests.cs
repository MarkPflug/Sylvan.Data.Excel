using System;
using System.IO;
using System.Threading.Tasks;
using Xunit;

namespace Sylvan.Data.Excel
{
	public
		class IOTests
	{
		public IOTests()
		{
			using (var file = File.Create("test.bin"))
			{
				file.SetLength(0x10000000);// 256MB
			}
		}

		[Theory]
		[InlineData(0x200)]
		[InlineData(0x1000)]
		[InlineData(0x10000)]
		[InlineData(0x100000)]
		public async Task ReadPerf(int size)
		{
			// experimenting with IO differences between net5.0 and net6.0
			using var stream = File.OpenRead("test.bin");
			await Process(stream, size);
		}

		static async Task Process(Stream s, int batchSize = 0x100000)
		{
			var buf = new byte[batchSize];
			for (int i = 0; ; i++)
			{
				var l = await s.ReadAsync(buf, 0, batchSize);
				if (l == 0) break;
			}
		}
	}
}
