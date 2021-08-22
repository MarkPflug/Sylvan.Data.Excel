using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Sylvan.Data.Excel
{
	partial class Ole2Package
	{
		public class Ole2Entry
		{
			enum EntryType
			{
				Invalid = 0,
				Storage = 1,
				Stream = 2,
				LockBytes = 3,
				Property = 4,
				Root = 5,
			}

			int entryIdx;

			uint childIdx;
			uint lIdx;
			uint rIdx;

			public Ole2Package Package { get; private set; }

			public String Name { get; private set; }

			static readonly Regex badChars = new Regex("\\W+");

			public Uri EntryUri
			{
				get
				{
					var safeName = badChars.Replace(Name, "");
					var uriStr = "/" + safeName;
					var uri = new Uri(uriStr, UriKind.Relative);					
					return uri;
				}
			}

			EntryType Type { get; set; }

			public uint StartSector { get; set; }
			public long StreamSize { get; set; }

			public IEnumerable<Ole2Entry> GetChildren()
			{

				if (childIdx == FreeSector) yield break;
				Ole2Entry child = Package.entryList[childIdx];
				foreach (Ole2Entry entry in child.EnumerateSiblings())
				{
					yield return entry;
				}
			}

			IEnumerable<Ole2Entry> EnumerateSiblings()
			{

				yield return this;

				if (lIdx != FreeSector)
					foreach (Ole2Entry entry in Package.entryList[lIdx].EnumerateSiblings())
						yield return entry;

				if (rIdx != FreeSector)
					foreach (Ole2Entry entry in Package.entryList[rIdx].EnumerateSiblings())
						yield return entry;
			}

			public Ole2Entry(Ole2Package package, Stream iStream, int entryIdx)
			{
				this.Package = package;
				this.entryIdx = entryIdx;

				BinaryReader reader = new BinaryReader(iStream, Encoding.Unicode);

				byte[] dirNameBytes = new byte[64];

				reader.Read(dirNameBytes, 0, 64);
				ushort nameLen = reader.ReadUInt16();

				this.Name = Encoding.Unicode.GetString(dirNameBytes, 0, nameLen);

				EntryType type = (EntryType)reader.ReadByte();
				byte color = reader.ReadByte();

				lIdx = reader.ReadUInt32();
				rIdx = reader.ReadUInt32();
				childIdx = reader.ReadUInt32();

				// skip the clsID
				reader.ReadInt32();
				reader.ReadInt32();
				reader.ReadInt32();
				reader.ReadInt32();

				uint state = reader.ReadUInt32();
				ulong createTime = reader.ReadUInt64();
				ulong modfiyTime = reader.ReadUInt64();

				this.StartSector = reader.ReadUInt32();
				this.StreamSize = reader.ReadInt64();
			}

			public override string ToString()
			{
				return string.Format("{0}: {1}, {2}, {3}", entryIdx, Name, Type, StreamSize);
			}

			public Ole2Stream Open()
			{
				var sectors = this.Package.GetStreamSectors(this.StartSector).ToArray();
				return new Ole2Stream(this.Package, sectors, StreamSize);
			}

		}
	}
}
