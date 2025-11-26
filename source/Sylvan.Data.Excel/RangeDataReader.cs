using System;
using System.Collections;
using System.Data.Common;
using System.Text.RegularExpressions;

namespace Sylvan.Data.Excel;

partial class RangeDataReader : DbDataReader
{
	readonly ExcelDataReader r;
	readonly string spec;

	int state = -2;
	int rowIdx = -1;
	int endRow = -1;
	int colOffset = -1;
	int fieldCount = -1;

	static readonly Regex RangeRegex = new Regex("^((?<sheet>[a-z0-9_]{1,40})|(\\'(?<sheet>[a-z0-9_() -]{1,40})\\'))\\!\\$([A-Z]{1,4})\\$(\\d{1,7})(\\:\\$([A-Z]{1,4})\\$(\\d{1,7}))?$", RegexOptions.IgnoreCase);

	public RangeDataReader(ExcelDataReader r, string spec)
	{
		this.r = r;
		this.spec = spec;

	}

	int Map(int idx)
	{
		if (idx < 0 || idx >= this.fieldCount)
			throw new ArgumentOutOfRangeException();

		return idx - this.colOffset;
	}

	public override object this[int ordinal] => this.GetValue(ordinal);

	public override object this[string name] => this.GetValue(GetOrdinal(name));

	public override int Depth => 1;

	public override int FieldCount => fieldCount;

	public override bool HasRows => throw new NotImplementedException();

	public override bool IsClosed => rowIdx == -3;

	public override int RecordsAffected => 0;

	public override bool GetBoolean(int ordinal)
	{
		return r.GetBoolean(Map(ordinal));
	}

	public override byte GetByte(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length)
	{
		throw new NotImplementedException();
	}

	public override char GetChar(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length)
	{
		throw new NotImplementedException();
	}

	public override string GetDataTypeName(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override DateTime GetDateTime(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override decimal GetDecimal(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override double GetDouble(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override IEnumerator GetEnumerator()
	{
		throw new NotImplementedException();
	}

	public override Type GetFieldType(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override float GetFloat(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override Guid GetGuid(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override short GetInt16(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override int GetInt32(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override long GetInt64(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override string GetName(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override int GetOrdinal(string name)
	{
		throw new NotImplementedException();
	}

	public override string GetString(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override object GetValue(int ordinal)
	{
		return r.GetValue(Map(ordinal));
	}

	public override int GetValues(object[] values)
	{
		var c = Math.Min(values.Length, this.FieldCount);
		for (int i = 0; i < c; i++)
		{
			values[i] = this.GetValue(i);
		}
		return c;
	}

	public override bool IsDBNull(int ordinal)
	{
		return r.IsDBNull(Map(ordinal));
	}

	public override bool NextResult()
	{
		return false;
	}

	void Init()
	{

		var qq = RangeRegex.Match(spec);
		if (qq.Success)
		{
			var sheet = qq.Groups["sheet"].Value;
			if (!r.TryOpenWorksheet(sheet))
			{
				throw new Exception();
			}

			string startColStr = qq.Groups[3].Value;
			int startCol = ParseCol(startColStr);
			string startRowStr = qq.Groups[4].Value;
			int startRow = int.Parse(startRowStr);

			string endColStr = qq.Groups[6].Value;
			int endCol = ParseCol(endColStr);
			string endRowStr = qq.Groups[7].Value;
			int endRow = int.Parse(endRowStr);

			this.colOffset = startCol;
			this.fieldCount = endCol - startCol + 1;
			this.endRow = endRow;
			this.rowIdx = 1;
			this.state = 1;
			// advance to the start of the range
			while (rowIdx < startRow)
			{
				rowIdx++;
				if (!r.Read())
				{
					this.state = -1;
					break;
				}
			}
		}
	}

	static int ParseCol(string str)
	{
		int col = -1;
		for (int i = 0; i < str.Length; i++)
		{
			var c = str[i];
			var v = c - 'A';
			if ((uint)v < 26u)
			{
				col = ((col + 1) * 26) + v;
			}
			else
			{
				break;
			}
		}
		return col;
	}

	public override bool Read()
	{
		if (this.state == -2)
		{
			Init();
		}


		if (this.rowIdx <= this.endRow)
		{
			this.rowIdx++;
			return r.Read();
		}

		return false;
	}
}
