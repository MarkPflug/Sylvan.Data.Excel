using System.Diagnostics;
using System.IO.Compression;

var file = "minimal.xlsx";
ZipFile.CreateFromDirectory("Content", file);

Process.Start(new ProcessStartInfo {  FileName = file, UseShellExecute = true});