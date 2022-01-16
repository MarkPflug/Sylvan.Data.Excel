using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace Sylvan.Data.Excel
{
	partial class Ole2Package
	{
		const ulong magicSig = 0xe11ab1a1e011cfd0;
		const ulong MaxSector = 0xfffffffa;
		const ulong DiFatSector = 0xfffffffc;
		const ulong FatSector = 0xfffffffd;
		const ulong EndOfChain = 0xfffffffe;
		const ulong FreeSector = 0xffffffff;
		const ushort ByteOrder = 0xfffe;
		const uint MiniSectorCutoff = 0x1000;
		const int HeaderFatSectorListCount = 109;
		const int DirectoryEntrySize = 0x80;

		BinaryReader reader;
		Stream stream;

		int sectorSize;
		ushort verMinor;
		ushort verMajor;
		uint directorySectorStart;
		uint directorySectorCount;
		uint fatSectorCount;
		uint miniSectorStart;
		uint miniFatSectorCount;
		uint miniSectorCutoff;


		uint fatSectorListStart;
		uint fatSectorListCount;
		uint[] fatSectorList;

		Ole2Entry[] entryList;

		public Ole2Entry RootEntry
		{
			get { return entryList[0]; }
		}

		public Ole2Entry GetEntry(int entryIdx)
		{
			return this.entryList[entryIdx];
		}

		public Ole2Entry? GetEntry(string name)
		{
			foreach (Ole2Entry entry in this.entryList)
				if (entry.Name == name)
					return entry;
			return null;
		}

		public Ole2Package(Stream iStream)
		{

			this.stream = iStream;
			this.reader = new BinaryReader(iStream, Encoding.Unicode);
			this.fatSectorList = Array.Empty<uint>();
			this.entryList = Array.Empty<Ole2Entry>();
			LoadHeader();
			LoadDirectoryEntries();
		}

		void LoadHeader()
		{

			BinaryReader reader = new BinaryReader(stream, Encoding.Unicode);

			ulong magic = reader.ReadUInt64();

			if (magic != magicSig)
				throw new InvalidDataException();//"Invalid file format"

			byte[] clsId = reader.ReadBytes(16);
			if (!clsId.All(b => b == 0))
				throw new InvalidDataException();//"Invalid class id"

			this.verMinor = reader.ReadUInt16();
			this.verMajor = reader.ReadUInt16();

			switch (verMajor)
			{
				case 3:
					this.sectorSize = 0x0200;
					break;
				case 4:
					this.sectorSize = 0x1000;
					break;
				default:
					throw new InvalidDataException();//"Invalid Ole2 version."
			}

			ushort byteOrder = reader.ReadUInt16();
			if (byteOrder != ByteOrder)
				throw new InvalidDataException();// "Invalid byte order marks."

			ushort sectorShift = reader.ReadUInt16();
			if ((1 << sectorShift) != sectorSize)
				throw new InvalidDataException();// "Invalid sector size"

			ushort miniSectorShift = reader.ReadUInt16();
			if (miniSectorShift != 6)
				throw new InvalidDataException();// "Invalid mini sector shift"

			reader.ReadUInt16(); // reserved
			reader.ReadUInt32(); // reserved

			this.directorySectorCount = reader.ReadUInt32();
			if (directorySectorCount != 0) // 0 for 512 sector
				throw new IOException();

			this.fatSectorCount = reader.ReadUInt32();

			directorySectorStart = reader.ReadUInt32();


			uint sig = reader.ReadUInt32();
			if (sig != 0) throw new InvalidDataException();// "Invalid transaction signature"


			this.miniSectorCutoff = reader.ReadUInt32();
			if (miniSectorCutoff != MiniSectorCutoff) throw new InvalidDataException();// "invalid mini sector cutoff"

			this.miniSectorStart = reader.ReadUInt32();
			this.miniFatSectorCount = reader.ReadUInt32();

			this.fatSectorListStart = reader.ReadUInt32();
			this.fatSectorListCount = reader.ReadUInt32();

			LoadFatSectorList();
		}

		void LoadDirectoryEntries()
		{
			int directorySectorCount = 0;

			foreach (uint v in GetStreamSectors(directorySectorStart))
				directorySectorCount++;

			int entryCount = sectorSize * directorySectorCount / DirectoryEntrySize;
			this.entryList = new Ole2Entry[entryCount];
			var sectors = GetStreamSectors(directorySectorStart).ToArray();
			Ole2Stream dirStream = new Ole2Stream(this, sectors, sectorSize * directorySectorCount);
			for (int i = 0; i < entryCount; i++)
			{
				this.entryList[i] = new Ole2Entry(this, dirStream, i);
			}
		}

		long SectorOffset(uint sectorIdx)
		{
			return (sectorIdx + 1) * sectorSize;
		}

		void LoadFatSectorList()
		{
			fatSectorList = new uint[HeaderFatSectorListCount + fatSectorListCount * (sectorSize / 4 - 1)];
			int i = 0;

			// load the FAT sectors list from the header sector
			for (; i < HeaderFatSectorListCount;)
			{
				fatSectorList[i++] = reader.ReadUInt32();
			}

			uint sect = fatSectorListStart;

			// load the FAT sectors list chained off the header
			for (int j = 0; j < this.fatSectorListCount; j++)
			{
				long sectOff = SectorOffset(sect);
				stream.Seek(sectOff, SeekOrigin.Begin);

				int diCount = sectorSize / 4 - 1;

				for (int k = 0; k < diCount; k++)
					fatSectorList[i++] = reader.ReadUInt32();

				// read next difat location
				sect = reader.ReadUInt32();
			}
		}

		IEnumerable<uint> GetStreamSectors(uint startSector)
		{
			uint sector = startSector;

			do
			{
				yield return sector;
				sector = NextSector(sector);
			} while (sector != EndOfChain);
		}

		public uint NextSector(uint sector)
		{
			uint fatSectIdx = (uint)(sector / (sectorSize / 4));
			uint fatSectOff = (uint)(sector % (sectorSize / 4));

			uint fatPage = fatSectorList[fatSectIdx];

			stream.Seek(SectorOffset(fatPage) + (fatSectOff * 4), SeekOrigin.Begin);
			return reader.ReadUInt32();
		}

		protected Ole2PackagePart[] GetPartsCore()
		{
			var entries = GetEntries();
			return entries.Select(entry => new Ole2PackagePart(entry)).ToArray();
		}

		IEnumerable<Ole2Entry> GetEntries()
		{
			var root = this.RootEntry;
			foreach (var entry in EnumerateEntry(root))
				yield return entry;
		}

		IEnumerable<Ole2Entry> EnumerateEntry(Ole2Entry entry)
		{
			yield return entry;
			foreach (var child in entry.GetChildren())
			{
				foreach (var descendant in EnumerateEntry(child))
				{
					yield return descendant;
				}
			}
		}
	}
}
