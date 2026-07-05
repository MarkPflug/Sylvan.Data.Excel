using System.Collections;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using Sylvan.Data.Excel;

// Round-trips string/int/enum columns through ExcelDataWriter -> ExcelDataReader
// under NativeAOT to prove the enum read/write paths (FieldAccessor.EnumAccessor<T>,
// XlsxDataWriter+FieldWriter.EnumFieldWriter) don't rely on reflection that NativeAOT
// can't satisfy (MakeGenericMethod / MakeGenericType / Activator.CreateInstance).

var rows = new[]
{
	(Id: 1, Name: "Alpha", Status: SmokeStatus.Active),
	(Id: 2, Name: "Bravo", Status: SmokeStatus.Inactive),
	(Id: 3, Name: "Charlie", Status: SmokeStatus.Pending),
};

using var ms = new MemoryStream();
using (var writer = ExcelDataWriter.Create(ms, ExcelWorkbookType.ExcelXml))
{
	writer.Write(new SmokeDataReader(rows));
}

ms.Position = 0;

using var reader = ExcelDataReader.Create(ms, ExcelWorkbookType.ExcelXml);
var actual = new List<(int Id, string Name, SmokeStatus Status)>();
while (reader.Read())
{
	actual.Add((
		reader.GetInt32(0),
		reader.GetString(1),
		reader.GetFieldValue<SmokeStatus>(2)
	));
}

if (actual.Count != rows.Length)
{
	Console.Error.WriteLine($"FAIL: expected {rows.Length} rows, got {actual.Count}");
	return 1;
}

for (var i = 0; i < rows.Length; i++)
{
	if (actual[i] != rows[i])
	{
		Console.Error.WriteLine($"FAIL: row {i} expected {rows[i]}, got {actual[i]}");
		return 1;
	}
}

Console.WriteLine($"OK: {actual.Count} rows round-tripped (string/int/enum) under NativeAOT.");
return 0;

enum SmokeStatus
{
	Active = 1,
	Inactive,
	Pending,
}

sealed class SmokeDataReader : DbDataReader
{
	readonly (int Id, string Name, SmokeStatus Status)[] rows;
	int row = -1;

	public SmokeDataReader((int Id, string Name, SmokeStatus Status)[] rows)
	{
		this.rows = rows;
	}

	public override int FieldCount => 3;

	public override string GetName(int ordinal) => ordinal switch
	{
		0 => "Id",
		1 => "Name",
		2 => "Status",
		_ => throw new ArgumentOutOfRangeException(nameof(ordinal)),
	};

	[return: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)]
	public override Type GetFieldType(int ordinal) => ordinal switch
	{
		0 => typeof(int),
		1 => typeof(string),
		2 => typeof(SmokeStatus),
		_ => throw new ArgumentOutOfRangeException(nameof(ordinal)),
	};

	public override int GetInt32(int ordinal) => rows[row].Id;

	public override string GetString(int ordinal) => rows[row].Name;

	public override object GetValue(int ordinal) => ordinal switch
	{
		0 => rows[row].Id,
		1 => rows[row].Name,
		2 => rows[row].Status,
		_ => throw new ArgumentOutOfRangeException(nameof(ordinal)),
	};

	public override bool IsDBNull(int ordinal) => false;

	public override bool Read()
	{
		row++;
		return row < rows.Length;
	}

	public override object this[int ordinal] => GetValue(ordinal);
	public override object this[string name] => throw new NotImplementedException();
	public override int Depth => 0;
	public override bool HasRows => rows.Length > 0;
	public override bool IsClosed => false;
	public override int RecordsAffected => -1;

	public override bool GetBoolean(int ordinal) => throw new NotImplementedException();
	public override byte GetByte(int ordinal) => throw new NotImplementedException();
	public override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length) => throw new NotImplementedException();
	public override char GetChar(int ordinal) => throw new NotImplementedException();
	public override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length) => throw new NotImplementedException();
	public override string GetDataTypeName(int ordinal) => GetFieldType(ordinal).Name;
	public override DateTime GetDateTime(int ordinal) => throw new NotImplementedException();
	public override decimal GetDecimal(int ordinal) => throw new NotImplementedException();
	public override double GetDouble(int ordinal) => throw new NotImplementedException();
	public override IEnumerator GetEnumerator() => throw new NotImplementedException();
	public override float GetFloat(int ordinal) => throw new NotImplementedException();
	public override Guid GetGuid(int ordinal) => throw new NotImplementedException();
	public override short GetInt16(int ordinal) => throw new NotImplementedException();
	public override long GetInt64(int ordinal) => throw new NotImplementedException();
	public override int GetOrdinal(string name) => throw new NotImplementedException();
	public override int GetValues(object[] values) => throw new NotImplementedException();
	public override bool NextResult() => false;
}
