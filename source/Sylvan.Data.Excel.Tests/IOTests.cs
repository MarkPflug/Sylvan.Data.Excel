using System;
using System.IO;
using System.Threading.Tasks;
using Xunit;

namespace Sylvan.Data.Excel
{
	public class IOTests
	{
		[Theory]
		[InlineData(1)]
		[InlineData(0x200)]
		[InlineData(0x1000)]
		[InlineData(0x10000)]
		[InlineData(0x100000)]
		public async Task ReadPerf(int size)
		{
			var file = "/data/excel/vehicles_xls.xls";
			//using var stream = File.OpenRead(file);
			using var stream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.Read, size);
			await Process(stream);
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
