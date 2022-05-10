using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Sylvan.Data.Excel;

partial class Ole2Package
{
	public sealed class Ole2Stream : Stream
	{
		Ole2Package package;

		readonly long length;
		long position;

		int sectorOff;
		readonly int sectorLen;

		readonly uint[] sectors;
		int sectorIdx;
		uint sector;

		long streamPos;

		public Ole2Stream(Ole2Package package, uint[] sectors, long length)
		{
			this.package = package;
			this.sectors = sectors;
			this.sectorIdx = 0;
			this.sector = sectors[sectorIdx];
			this.length = length;
			this.position = 0;
			this.sectorLen = this.package.sectorSize;
			this.sectorOff = 0;
			this.streamPos = -1;
		}

		//bool NextSector()
		//{
		//	sectorIdx++;
		//	if (sectorIdx < sectors.Length)
		//	{
		//		this.sector = sectors[sectorIdx];

		//		this.sectorOff = 0;
		//		return true;
		//	}
		//	return false;
		//}

		public override long Position
		{
			get
			{
				return this.position;
			}
			set
			{
				Seek(value - this.position, SeekOrigin.Current);
			}
		}

		public override long Seek(long offset, SeekOrigin origin)
		{
			throw new NotSupportedException();
		}

		public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
		{
			if(offset + count > buffer.Length)
				throw new ArgumentOutOfRangeException();

			//Debug.WriteLine($"{offset} {count} {this.position}");
						
			var sectors = this.sectors;

			int bytesRead = 0;
			var c = count;

			while (bytesRead < count && position < length)
			{
				var readLen = 0;
				var readStart = (sector + 1) * sectorLen + sectorOff;
				var curSector = sector;
				
				while (readLen < c)
				{
					if (this.sectorOff >= this.sectorLen)
					{
						sectorOff = 0;
						sectorIdx++;
						if (sectorIdx >= sectors.Length)
						{
							break;
						}
						var nextSector = sectors[sectorIdx];
						if (nextSector != curSector + 1)
						{
							// next sector is not coniguious, so read
							// the current contig block
							sector = nextSector;
							if(readLen > 0) { 
								break;
							}
						}
						sector = curSector = nextSector;
					}

					var sectorAvail = this.sectorLen - this.sectorOff;
					var sectorRead = Math.Min(sectorAvail, c - readLen);

					readLen += sectorRead;
					this.sectorOff += sectorRead;
				}

				// avoid seek if we are already positioned.
				if (streamPos != readStart)
				{
					package.stream.Seek(readStart, SeekOrigin.Begin);
					streamPos = readStart;
				}

				if (readLen == 0)
					break;
				int len = 0;
				while (len < readLen)
				{
					int l = await package.stream.ReadAsync(buffer, offset, readLen);
					if (l == 0)
						throw new IOException();//"Unexpectedly encountered end of Ole2Package Stream"
					len += l;
					offset += l;
					c -= l;
					this.position += l;
					this.streamPos += l;
				}
				bytesRead += len;
			}
			return bytesRead;
		}

		public override int Read(byte[] buffer, int offset, int count)
		{
			return ReadAsync(buffer, offset, count, default).GetAwaiter().GetResult();
		}

		public override void Flush()
		{
		}

		public override void SetLength(long value)
		{
			throw new NotSupportedException();
		}

		public override void Write(byte[] buffer, int offset, int count)
		{
			throw new NotSupportedException();
		}

		public override bool CanRead => true;

		public override bool CanSeek => true;

		public override bool CanWrite => false;

		public override long Length => length;
	}
}
