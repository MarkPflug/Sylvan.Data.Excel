using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using static Sylvan.Data.Excel.Ole2Package;

namespace Sylvan.Data.Excel.Xls;

sealed partial class XlsWorkbookReader
{
	sealed class RecordReader
	{
		const int BufferSize = 0x40000;
		const int MaxRecordSize = 8228;

		Ole2Stream stream;

		// the working buffer. The current biff record is guaranteed to be loaded entirely in this buffer.
		byte[] buffer;
		// the length of the data in the current buffer.
		int bufferLen;
		// the current position in the working buffer.
		int bufferPos;

		// buffer used to assemble large strings.
		char[] strBuffer;

		short recordCode;
		// the offset of the start of the current record in the buffer
		int recordOff = 0;
		// the length of the current record.
		int recordLen;

		public RecordType Type { get { return (RecordType)recordCode; } }
		public int Length { get { return recordLen; } }

		public RecordReader(Ole2Stream stream)
		{
			this.stream = stream;
			this.buffer = new byte[BufferSize];
			this.bufferLen = 0;
			this.bufferPos = 0;
			this.strBuffer = Array.Empty<char>();
		}

		async Task<bool> FillBufferAsync(int required)
		{
			var len = bufferLen - recordOff;

			if (len > 0)
			{
				Buffer.BlockCopy(buffer, recordOff, buffer, 0, len);
			}

			var shift = bufferLen - len;
			this.recordOff -= shift;
			this.bufferPos -= shift;
			this.bufferLen = len;

			Debug.Assert(recordOff == 0);
			Assert();
			int c = 0;

			while (c < required)
			{
				var l = await stream.ReadAsync(buffer, len, BufferSize - len, default).ConfigureAwait(false);
				c += l;
				if (l == 0)
				{
					break;
				}
				this.bufferLen = len + c;
			}
			return c >= required;
		}

		public void SetPosition(long offset)
		{
			this.stream.Seek(offset, SeekOrigin.Begin);
			this.bufferLen = 0;
			this.bufferPos = 0;
			this.recordOff = 0;
			this.recordLen = 0;
		}

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public byte ReadByte()
		{
			// the byte we are reading must be within the current record.
			Debug.Assert(bufferPos < recordOff + recordLen);
			Assert();
			var b = buffer[bufferPos];
			bufferPos++;
			return b;
		}

		public ushort PeekRow()
		{
			return BitConverter.ToUInt16(buffer, bufferPos);
		}

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public short ReadInt16()
		{
			return (short)(ReadByte() | ReadByte() << 8);
		}

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public ushort ReadUInt16()
		{
			return (ushort)(ReadByte() | ReadByte() << 8);
		}

		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public int ReadInt32()
		{
			return ReadByte() | ReadByte() << 8 | ReadByte() << 16 | ReadByte() << 24;
		}

		public async Task<string> ReadString16()
		{
			if (bufferPos >= recordOff + recordLen)
			{
				var next = await NextRecordAsync().ConfigureAwait(false);
				if (!next || Type != RecordType.Continue)
					throw new InvalidDataException();
			}

			// the length of the string in *characters*
			int len = ReadInt16();
			byte options = ReadByte();

			bool compressed = (options & 0x01) == 0;
			bool asian = (options & 0x04) != 0;
			bool rich = (options & 0x08) != 0;

			int richCount = 0;
			if (rich)
				richCount = ReadInt16();

			int asianCount = 0;
			if (asian)
				asianCount = ReadInt32();

			var str = await ReadStringBufferAsync(len, compressed).ConfigureAwait(false);

			var remain = richCount * 4 + asianCount;

			while (remain > 0)
			{
				var avail = recordOff + recordLen - bufferPos;
				var c = Math.Min(remain, avail);
				remain -= c;
				bufferPos += c;
				Assert();
				if (remain > 0)
				{
					var next = await NextRecordAsync().ConfigureAwait(false);
					if (!next || Type != RecordType.Continue)
						throw new InvalidDataException();
				}
			}

			return str;
		}

		static readonly Encoding Encoding1252 = Encoding.GetEncoding(1252);

		internal async Task<string> ReadStringBufferAsync(int charCount, bool compressed)
		{
			var strLen = charCount;
			// stores our position in the string we are assembling.
			int strPos = 0;
			for (int i = 0; ; i++)
			{
				var encoding = compressed ? Encoding1252 : Encoding.Unicode;
				int charSize = compressed ? 1 : 2;
				int byteCount = charCount * charSize;

				var recordPos = bufferPos - recordOff;
				int recordBytes = recordLen - recordPos;

				// if the string sits entirely within the current record
				// we can directly create a string from it.
				if (i == 0 && recordBytes >= byteCount)
				{
					var str = encoding.GetString(buffer, bufferPos, byteCount);
					bufferPos += byteCount;
					Assert();
					return str;
				}

				// bump up the buffer to hold the data if it's not big enough already.
				if (strLen > strBuffer.Length)
				{
					// TODO: should add a little headroom here. I'm finding this gets repeatedly resized
					// as ever-larger strings are encountered.
					Array.Resize(ref strBuffer, strLen);
				}

				// one of the following needs to be true
				// uncompressed string
				// the bytes in the string are all contained in the current record
				// the string overflows into the next record, and this current record contains an even number of bytes.
				Debug.Assert(charSize == 1 || byteCount < recordBytes || (recordBytes & 0x01) == 0);

				int currentRecordBytes = Math.Min(byteCount, recordBytes);
				var c = encoding.GetChars(buffer, bufferPos, currentRecordBytes, strBuffer, strPos);
				bufferPos += currentRecordBytes;
				Assert();

				charCount -= c;
				strPos += c;

				if (charCount > 0)
				{
					var next = await NextRecordAsync().ConfigureAwait(false);
					if (!next || Type != RecordType.Continue)
						throw new InvalidDataException();

					var b = ReadByte();
					compressed = b == 0;
					continue;
				}
				else
				{
					break;
				}
			}
			return new string(strBuffer, 0, strLen);
		}

		public async Task<string> ReadByteString(int lenSize)
		{
			int len;
			if (lenSize == 1)
				len = ReadByte();
			else
				len = ReadInt16();

			await ReadStringBufferAsync(len, true).ConfigureAwait(false);
			var str = new string(strBuffer, 0, len);
			return str;
		}

		public async Task<string> ReadString8()
		{
			int len = ReadByte();
			byte options = ReadByte();

			bool compressed = (options & 0x01) == 0;
			bool asian = (options & 0x04) != 0;
			bool rich = (options & 0x08) != 0;

			int richCount = 0;
			if (rich)
				richCount = ReadInt16();

			int asianCount = 0;
			if (asian)
				asianCount = ReadInt32();

			var str = await ReadStringBufferAsync(len, compressed).ConfigureAwait(false);

			for (int i = 0; i < richCount; i++)
			{
				ReadInt32();
			}

			for (int i = 0; i < asianCount; i++)
			{
				ReadByte();
			}

			return str;
		}

		public async Task<string> ReadStringAsync(int length, bool compressed)
		{
			var str = await ReadStringBufferAsync(length, compressed).ConfigureAwait(false);
			return str;
		}

		// reads the next BIFF record. Ensuring the entire
		// record bytes are in the working buffer.
		public async Task<bool> NextRecordAsync()
		{
			bufferPos = recordOff + recordLen;

			if (bufferPos + 4 >= bufferLen)
			{
				if (!await FillBufferAsync(4).ConfigureAwait(false))
				{
					return false;
				}
			}
			this.recordOff = bufferPos;
			this.recordLen = 4; // we have at least the first 4 bytes.
			this.recordCode = ReadInt16();
			Debug.Assert(recordCode >= 0);
			this.recordLen = ReadInt16();
			if (recordLen < 0 || recordLen > MaxRecordSize)
				throw new InvalidDataException();
			Debug.Assert(recordLen >= 0);

			this.recordOff = bufferPos;
			if (recordOff + recordLen > bufferLen)
			{
				var req = (recordOff + recordLen) - bufferLen;
				Debug.Assert(req >= 1);

				if (!await FillBufferAsync(req).ConfigureAwait(false))
				{
					return false;
				}
			}

			//Debug.WriteLine($"{(RecordType)this.recordCode} {this.recordCode:x} {this.recordLen}");
			return true;
		}

		[Conditional("DEBUG")]
		void Assert()
		{
			Debug.Assert(bufferPos >= 0 && bufferPos <= bufferLen);
		}
	}
}
