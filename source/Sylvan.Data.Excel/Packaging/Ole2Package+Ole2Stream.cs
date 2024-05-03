using System;
using System.IO;

namespace Sylvan.Data.Excel;

partial class Ole2Package
{
	public sealed class Ole2StreamOld : Stream
	{
		readonly Stream stream;

		readonly long length;
		readonly int sectorLen;
		readonly int startSector;
		readonly uint[] sectors;
		
		long position;
		int sectorOff;

		int sectorIdx;
		uint sector;

		public Ole2StreamOld(Stream stream, uint[] sectors, int sectorLen, int startSector, long length)
		{
			this.stream = stream;
			this.sectors = sectors;
			this.sectorIdx = 0;
			this.sector = sectors[sectorIdx];
			this.startSector = startSector;
			this.length = length;
			this.position = 0;
			this.sectorLen = sectorLen;
			this.sectorOff = 0;
		}

		public override long Position
		{
			get
			{
				return this.position;
			}
			set
			{
				Seek(value, SeekOrigin.Begin);
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
			if (pos < 0 || pos > this.length)
			{
				throw new ArgumentOutOfRangeException(nameof(offset));
			}

			this.position = pos;
			var idx = pos / this.sectorLen;

			this.sectorIdx = (int)idx;
			this.sectorOff = (int)(pos - (idx * sectorLen));
			this.sector = this.sectors[sectorIdx];
			return this.position;
		}

		public override int Read(byte[] buffer, int offset, int count)
		{
			if (offset + count > buffer.Length)
				throw new ArgumentOutOfRangeException();

			var stream = this.stream;

			var sectors = this.sectors;

			int bytesRead = 0;
			var c = count;

			while (bytesRead < count && position < length)
			{
				var readLen = 0;
				var readStart = (sector + startSector) * sectorLen + sectorOff;
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
							if (readLen > 0)
							{
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
				if (stream.Position != readStart)
				{
					stream.Seek(readStart, SeekOrigin.Begin);
				}

				if (readLen == 0)
					break;
				int len = 0;
				while (len < readLen)
				{
					int l = stream.Read(buffer, offset, readLen);
					if (l == 0)
					{
						throw new IOException();
					}
					len += l;
					offset += l;
					c -= l;
					this.position += l;
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
