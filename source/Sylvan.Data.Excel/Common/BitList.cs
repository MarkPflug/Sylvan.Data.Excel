using System;

namespace Sylvan;

// this is used to track hidden columns.
// hiding columns is rare, so this is designed
// to not allocate in the case that there are no hidden columns (no true bits)
sealed class BitList
{
	int[] data;

	public BitList()
	{
		this.data = Array.Empty<int>();
	}

	public bool this[int idx]
	{
		get
		{
			var q = idx / 32;
			var r = idx - q * 32;

			if (q >= data.Length) return false;

			var v = data[q];
			return (v & (1 << r)) != 0;
		}
		set
		{
			if (value == false)
			{
				// default is false, so we don't need to store anything
				return;
			}

			var q = idx / 32;
			var r = idx - q * 32;

			if (q >= data.Length)
			{
				if (value == false)
				{
					return;
				}
				else
				{
					// TODO: should this just use the normal x2 growth
					// strategy instead?
					Array.Resize(ref data, q + 1);
				}
			}

			var v = data[q];
			data[q] = v | 1 << r;
		}
	}
}