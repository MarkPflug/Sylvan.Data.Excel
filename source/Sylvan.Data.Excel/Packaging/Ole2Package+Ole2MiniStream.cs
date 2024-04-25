using System;
using System.Diagnostics;
using System.IO;

namespace Sylvan.Data.Excel;

partial class Ole2Package
{
	public sealed class Ole2MiniStream : Stream
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
		Stream miniStream;

		public Ole2MiniStream(Ole2Package package, Stream miniStream, uint[] sectors, long length)
		{
			this.miniStream = miniStream;
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
			long pos = 0;
			switch (origin)
			{
				case SeekOrigin.Begin:
					pos = offset;
					break;
				case SeekOrigin.Current:
					pos = this.position + offset;
					break;
				case SeekOrigin.End:
					pos = this.length + offset;
					break;
			}
			if (pos < 0)
			{
				throw new ArgumentOutOfRangeException(nameof(offset));
			}

			this.position = pos;
			var idx = pos / MiniSectorSize;

			this.sectorIdx = (int)idx;
			this.sectorOff = (int) (pos - (idx * MiniSectorSize));
			this.sector = this.sectors[sectorIdx];
			return this.position;
		}

		public override int Read(byte[] buffer, int offset, int count)
		{
			if (offset + count > buffer.Length)
				throw new ArgumentOutOfRangeException();

			var sectors = this.sectors;

			int bytesRead = 0;
			var c = count;
			var z = (int) Math.Min(count, this.length - position);

			while (bytesRead < count && position < length)
			{
				var readLen = 0;
				var readStart = sector * MiniSectorSize + sectorOff;
				var curSector = sector;

				while (readLen < z)
				{
					if (this.sectorOff >= MiniSectorSize)
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
							if (readLen > 0)
							{
								break;
							}
						}
						sector = curSector = nextSector;
					}

					var sectorAvail = MiniSectorSize - this.sectorOff;
					Debug.WriteLine("SA: " + sectorAvail);
					var sectorRead = Math.Min(sectorAvail, z - readLen);

					readLen += sectorRead;
					this.sectorOff += sectorRead;
				}

				// avoid seek if we are already positioned.
				if (streamPos != readStart)
				{
					package.miniStream.Seek(readStart, SeekOrigin.Begin);
					streamPos = readStart;
				}

				if (readLen == 0)
					break;
				int len = 0;
				while (len < readLen)
				{
					int l = package.miniStream.Read(buffer, offset, readLen);
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
