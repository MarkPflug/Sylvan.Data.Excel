using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Sylvan.Data.Excel
{
	sealed partial class XlsWorkbookReader
	{
		sealed class RecordReader
		{
			const int BufferSize = 0x40000;

			Stream stream;

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

			public RecordReader(Stream stream)
			{
				this.stream = stream;
				this.buffer = new byte[BufferSize];
				this.bufferLen = 0;
				this.bufferPos = 0;
				this.strBuffer = Array.Empty<char>();
			}

			async Task FillBufferAsync()
			{
				var len = bufferLen - recordOff;
				Buffer.BlockCopy(buffer, recordOff, buffer, 0, len);

				var shift = bufferLen - len;
				recordOff -= shift;
				bufferPos -= shift;

				var c = await stream.ReadAsync(buffer, len, BufferSize - len, default);
				this.bufferLen = len + c;
			}

			[MethodImpl(MethodImplOptions.AggressiveInlining)]
			public byte ReadByte()
			{
				// the byte we are reading must be within the current record.
				Debug.Assert(bufferPos < recordOff + recordLen);
				var b = buffer[bufferPos];
				bufferPos++;
				return b;
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
					var next = await NextRecordAsync();
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

				var str = await ReadStringBufferAsync(len, compressed);

				for (int i = 0; i < richCount; i++)
				{
					short offset = ReadInt16();
					short fontIdx = ReadInt16();
				}

				for (int i = 0; i < asianCount; i++)
				{
					ReadByte();
				}

				return str;
			}

			internal async Task<string> ReadStringBufferAsync(int charCount, bool compressed)
			{
				var strLen = charCount;
				// stores our position in the string we are assembling.
				int strPos = 0;
				for (int i = 0; ; i++)
				{
					var encoding = compressed ? Encoding.ASCII : Encoding.Unicode;
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

					charCount -= c;
					strPos += c;

					if (charCount > 0)
					{
						var next = await NextRecordAsync();
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

				await ReadStringBufferAsync(len, true);
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

				var str = await ReadStringBufferAsync(len, compressed);

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
				Debug.WriteLine("ReadString");
				var str = await ReadStringBufferAsync(length, compressed);
				return str;
			}

			// reads the next BIFF record. Ensuring the entire
			// record bytes are in the working buffer.
			public async Task<bool> NextRecordAsync()
			{
				bufferPos = recordOff + recordLen;

				if (bufferPos + 4 >= bufferLen)
				{
					await FillBufferAsync();
				}
				this.recordOff = bufferPos;
				this.recordLen = 4; // we have at least the first 4 bytes.
				this.recordCode = ReadInt16();
				this.recordLen = ReadInt16();

				if (bufferPos + recordLen >= bufferLen)
				{
					await FillBufferAsync();
				}

				this.recordOff = bufferPos;

				return true;
			}
		}
	}
}
