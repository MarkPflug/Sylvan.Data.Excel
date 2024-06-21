using System;
using System.Diagnostics;
using System.IO;

namespace Sylvan.Data.Excel;

partial class Ole2Package
{
	public sealed class Ole2Stream : Stream
	{
		readonly Stream stream;
		readonly long length;
		readonly int sectorLen;
		readonly int startSector; // either 0 or 1 depending on if this is a ministream or not
		readonly uint[] sectors;
		
		long position;

		public Ole2Stream(Stream stream, uint[] sectors, int sectorLen, int startSector, long length)
		{
			Debug.Assert(startSector == 0 || startSector == 1);

			this.stream = stream;
			this.sectors = sectors;
			this.startSector = startSector;
			this.length = length;
			this.position = 0;
			this.sectorLen = sectorLen;
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

		static int DivRem(int n, int d, out int r)
		{
			var q = n / d;

			r = n - q * d;

			return q;
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

			return this.position;
		}

		public override int Read(byte[] buffer, int offset, int count)
		{
			if (offset + count > buffer.Length)
				throw new ArgumentOutOfRangeException();

			var stream = this.stream;

			var sectors = this.sectors;

			var pos = this.position;

			int bytesRead = 0;
	
			var streamAvail = this.length - pos;

			// the amount to read is the lesser of the user requested count
			// and what remains in the stream.
			var readAvail = (int)Math.Min(count, streamAvail);
			
			var readRemain = readAvail;

			while (bytesRead < readAvail)
			{				
				// determine the longest block that can be read
				// in a single IO request, as sectors are often contiguous.
				var readLen = 0;
				
				if (pos > int.MaxValue) throw new NotSupportedException(); // TODO

				int sectorOff;
				var sectorIdx = DivRem((int)pos, sectorLen, out sectorOff);
				var sector = this.sectors[sectorIdx];
				
				var readStart = (startSector + sector) * sectorLen + sectorOff;

				while (readLen < readRemain)
				{
					if (sectorOff >= this.sectorLen)
					{
						Debug.Assert(sectorOff == this.sectorLen);
						sectorOff = 0;
						sectorIdx++;
						if (sectorIdx >= sectors.Length)
						{
							break;
						}
						var nextSector = sectors[sectorIdx];
						if (nextSector != sector + 1)
						{
							// next sector is not coniguious, so read
							// the current contig block
							//sector = nextSector;
							if (readLen > 0)
							{
								break;
							}
						}
						sector = nextSector;
					}

					var sectorAvail = this.sectorLen - sectorOff;
					var sectorRead = Math.Min(sectorAvail, readRemain - readLen);

					readLen += sectorRead;
					sectorOff += sectorRead;
				}
								// avoid seek if we are already positioned.
				if (stream.Position != readStart)
				{
					stream.Seek(readStart, SeekOrigin.Begin);
				}

				Debug.Assert(pos + readLen <= length);

				readRemain -= readLen;

				while (readLen > 0)
				{
					int l = stream.Read(buffer, offset, readLen);
					if (l == 0)
					{
						throw new IOException();
					}
					readLen -= l;
					offset += l;
					bytesRead += l;
					pos += l;					
				}
			}
			this.position = pos;
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
