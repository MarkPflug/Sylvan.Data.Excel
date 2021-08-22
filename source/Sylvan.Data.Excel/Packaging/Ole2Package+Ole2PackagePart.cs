using System;
using System.IO;
using System.Linq;

namespace Sylvan.Data.Excel
{
	partial class Ole2Package
	{
		public sealed class Ole2PackagePart
		{
			Ole2Entry entry;

			public Ole2PackagePart(Ole2Entry entry)
			{
				this.entry = entry;
			}

			public Stream GetStream()
			{
				var sectors = entry.Package.GetStreamSectors(entry.StartSector).ToArray();
				return new Ole2Stream(entry.Package, sectors, entry.StreamSize);
			}
		}
	}
}
