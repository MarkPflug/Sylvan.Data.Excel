using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sylvan.Testing
{
	class TestStream : Stream
	{
		Stream stream;
		public TestStream(Stream stream)
		{
			this.stream = stream;
		}

		public override bool CanRead => this.stream.CanRead;

		public override bool CanSeek => this.stream.CanSeek;

		public override bool CanWrite => this.stream.CanWrite;

		public override long Length => this.stream.Length;

		public bool IsClosed { get; private set; }

		public override void Close()
		{
			this.IsClosed = true;
			base.Close();
		}

		public override long Position { 
			get => this.stream.Position; 
			set => this.stream.Position = value; 
		}

		public override void Flush()
		{
			this.stream.Flush();
		}

		public override int Read(byte[] buffer, int offset, int count)
		{
			return this.stream.Read(buffer, offset, count);
		}

		public override long Seek(long offset, SeekOrigin origin)
		{
			return this.stream.Seek(offset, origin);
		}

		public override void SetLength(long value)
		{
			this.stream.SetLength(value);
		}

		public override void Write(byte[] buffer, int offset, int count)
		{
			this.stream.Write(buffer, offset, count);
		}
	}
}