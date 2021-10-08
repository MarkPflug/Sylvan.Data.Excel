using System;
using System.Data;
using System.Data.Common;
using System.Linq;
using System.Threading.Tasks;

namespace Sylvan.Data
{
	static class DataExtensions
	{
		public static void ProcessStrings(this IDataReader reader)
		{
			while (reader.Read())
			{
				for (int i = 0; i < reader.FieldCount; i++)
				{
					reader.GetString(i);
				}
			}
		}

		public static void ProcessValues(this IDataReader reader)
		{
			while (reader.Read())
			{
				for (int i = 0; i < reader.FieldCount; i++)
				{
					reader.GetValue(i);
				}
			}
		}

		public static void Process(this IDataReader reader)
		{
			TypeCode[] types = new TypeCode[reader.FieldCount];
			for (int i = 0; i < reader.FieldCount; i++)
			{
				var t = reader.GetFieldType(i);
				t = Nullable.GetUnderlyingType(t) ?? t;
				types[i] = Type.GetTypeCode(t);
			}

			while (reader.Read())
			{
				for (int i = 0; i < reader.FieldCount; i++)
				{
					if (reader.IsDBNull(i))
						continue;
					ProcessField(reader, i, types[i]);
				}
			}
		}

		public static void Process(this DbDataReader reader)
		{
			var cols = reader.GetColumnSchema();
			bool[] allowDbNull = cols.Select(c => c.AllowDBNull != false).ToArray();
			TypeCode[] types = cols.Select(c => Type.GetTypeCode(c.DataType)).ToArray();
			while (reader.Read())
			{
				for (int i = 0; i < reader.FieldCount; i++)
				{
					if (allowDbNull[i] && reader.IsDBNull(i))
						continue;
					ProcessField(reader, i, types[i]);
				}
			}
		}

		public static async Task ProcessAsync(this DbDataReader reader)
		{
			var cols = reader.GetColumnSchema();
			bool[] allowDbNull = cols.Select(c => c.AllowDBNull != false).ToArray();
			TypeCode[] types = cols.Select(c => Type.GetTypeCode(c.DataType)).ToArray();
			while (await reader.ReadAsync())
			{
				for (int i = 0; i < reader.FieldCount; i++)
				{
					if (allowDbNull[i] && await reader.IsDBNullAsync(i))
						continue;

					ProcessField(reader, i, types[i]);
				}
			}
		}

		static void ProcessField(this IDataReader reader, int i, TypeCode typeCode)
		{
			switch (typeCode)
			{
				case TypeCode.Boolean:
					reader.GetBoolean(i);
					break;
				case TypeCode.Int32:
					reader.GetInt32(i);
					break;
				case TypeCode.DateTime:
					reader.GetDateTime(i);
					break;
				case TypeCode.Single:
					reader.GetFloat(i);
					break;
				case TypeCode.Double:
					reader.GetDouble(i);
					break;
				case TypeCode.Decimal:
					reader.GetDecimal(i);
					break;
				case TypeCode.String:
					reader.GetString(i);
					break;
				default:
					// no cheating
					throw new NotSupportedException("" + typeCode);
			}
		}
	}
}
