sealed class XlsbReader
{
	byte[] data;
	int idx = 0;

	public XlsbReader(Stream stream)
	{
		var ms = new MemoryStream();
		stream.CopyTo(ms);
		data = ms.ToArray();
	}

	public ReadOnlySpan<byte> RecordSpan
	{
		get
		{
			return data.AsSpan(start, end - start);
		}
	}

	public ReadOnlySpan<byte> DataSpan
	{
		get
		{
			return data.AsSpan(dataStart, end - dataStart);
		}
	}

	public int Type => type;
	public int Length => len;

	int start;
	int dataStart;
	int end;
	int type;
	int len;

	public bool ReadRecord()
	{
		if (idx >= data.Length)
			return false;

		this.start = idx;

		var i = idx;

		this.type = ReadRecordType(ref i);
		this.len = ReadRecordLen(ref i);
		this.dataStart = i;

		i += len;

		this.end = i;
		this.idx = i;
		return true;
	}

	int ReadRecordType(ref int idx)
	{
		var b = data[idx++];
		int type;
		if (b >= 0x80)
		{
			var b2 = data[idx++];
			if (b2 >= 0x80)
				throw new InvalidDataException();
			type = (b & 0x7f) | (b2 << 7);
		}
		else
		{
			type = b;
		}
		return type;
	}

	int ReadRecordLen(ref int idx)
	{
		int accum = 0;
		int shift = 0;
		for (int i = 0; i < 4; i++, shift += 7)
		{
			var b = data[idx++];
			accum |= (b & 0x7f) << shift;
			if ((b & 0x80) == 0)
				break;
		}
		return accum;
	}
}
