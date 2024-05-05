using System.IO;

namespace Sylvan.Testing;

sealed class TestStream : Stream
{
	readonly Stream stream;

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
		stream.Close();
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