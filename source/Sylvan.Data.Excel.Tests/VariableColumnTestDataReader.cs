using System;
using System.Collections;
using System.Data.Common;

namespace Sylvan.Data.Excel;

// allows specifying a number of columns in the data reader for testing.
sealed class VariableColumnTestDataReader : DbDataReader
{
	int cols;
	int rows;
	int idx;

	public VariableColumnTestDataReader(int cols, int rows = 1)
	{
		this.cols = cols;
		this.rows = rows;
	}

	public override object this[int ordinal] => this.GetValue(ordinal);

	public override object this[string name] => this.GetValue(this.GetOrdinal(name));

	public override int Depth => 0;

	public override int FieldCount => cols;

	public override bool HasRows => true;

	public override bool IsClosed => false;

	public override int RecordsAffected => 1;

	public override bool GetBoolean(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override byte GetByte(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override long GetBytes(int ordinal, long dataOffset, byte[] buffer, int bufferOffset, int length)
	{
		throw new NotImplementedException();
	}

	public override char GetChar(int ordinal)
	{
		throw new NotImplementedException();
	}

	public override long GetChars(int ordinal, long dataOffset, char[] buffer, int bufferOffset, int length)
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
		return typeof(string);
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
		return "Col" + ordinal;
	}

	public override int GetOrdinal(string name)
	{
		throw new NotImplementedException();
	}

	public override string GetString(int ordinal)
	{
		return "Value" + ordinal;
	}

	public override object GetValue(int ordinal)
	{
		return GetString(ordinal);
	}

	public override int GetValues(object[] values)
	{
		throw new NotImplementedException();
	}

	public override bool IsDBNull(int ordinal)
	{
		return false;
	}

	public override bool NextResult()
	{
		return false;
	}

	public override bool Read()
	{
		return idx++ < rows;
	}
}
