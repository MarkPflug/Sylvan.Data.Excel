using System.Diagnostics;
using System.IO.Compression;

var file = "minimal.xlsx";
if (File.Exists(file))
	File.Delete(file);

ZipFile.CreateFromDirectory("Content", file);

Process.Start(new ProcessStartInfo { FileName = file, UseShellExecute = true });