using System;
using System.Collections.Generic;
using System.IO;

namespace Sylvan.Data.Excel;

/// <summary>
/// Exposes information about the supported Excel file types.
/// </summary>
public sealed class ExcelFileType
{
	/// <summary>
	/// The file extension for .xls files.
	/// </summary>
	public const string ExcelFileExtension = ".xls";

	/// <summary>
	/// The content type for .xls files.
	/// </summary>
	public const string ExcelContentType = "application/vnd.ms-excel";

	/// <summary>
	/// The file extension for .xlsx files.
	/// </summary>
	public const string ExcelXmlFileExtension = ".xlsx";

	/// <summary>
	/// The content type for .xlsx files.
	/// </summary>
	public const string ExcelXmlContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

	/// <summary>
	/// The file extension for .xlsm files.
	/// </summary>
	public const string ExcelXmlMacroEnabledFileExtension = ".xlsm";

	/// <summary>
	/// The content type for .xlsx files.
	/// </summary>
	public const string ExcelXmlMacroEnabledContentType = "application/vnd.ms-excel.sheet.macroEnabled.12";

	/// <summary>
	/// The file extension for .xlsb files.
	/// </summary>
	public const string ExcelBinaryFileExtension = ".xlsb";

	/// <summary>
	/// The content type for .xlsx files.
	/// </summary>
	public const string ExcelBinaryContentType = "application/vnd.ms-excel.sheet.binary.macroEnabled.12";


	/// <summary>
	/// The .xls file type.
	/// </summary>
	public static readonly ExcelFileType Excel =
		new(ExcelFileExtension,
			ExcelContentType,
			ExcelWorkbookType.Excel
		);

	/// <summary>
	/// The .xlsx file type.
	/// </summary>
	public static readonly ExcelFileType ExcelXml =
		new(
			ExcelXmlFileExtension,
			ExcelXmlContentType,
			ExcelWorkbookType.ExcelXml
		);

	/// <summary>
	/// The .xlsm file type.
	/// </summary>
	public static readonly ExcelFileType ExcelXmlMacroEnabled =
		new(
			ExcelXmlMacroEnabledFileExtension,
			ExcelXmlMacroEnabledContentType,
			ExcelWorkbookType.ExcelXml
		);

	/// <summary>
	/// The .xlsb file type.
	/// </summary>
	public static readonly ExcelFileType ExcelBinary =
		new(
			ExcelBinaryFileExtension,
			ExcelBinaryContentType,
			ExcelWorkbookType.ExcelBinary
		);

	/// <summary>
	/// Enumerates all the file types exposed by the Sylvan library.
	/// </summary>
	public static IEnumerable<ExcelFileType> All
	{
		get
		{
			yield return Excel;
			yield return ExcelXml;
			yield return ExcelXmlMacroEnabled;
			yield return ExcelBinary;
		}
	}

	/// <summary>
	/// Enumerates all file types supported by the ExcelDataReader.
	/// </summary>
	public static IEnumerable<ExcelFileType> ReaderSupported
	{
		get
		{
			yield return Excel;
			yield return ExcelXml;
			yield return ExcelXmlMacroEnabled;
			yield return ExcelBinary;
		}
	}

	/// <summary>
	/// Enumerates all file types supported by the ExcelDataWriter.
	/// </summary>
	public static IEnumerable<ExcelFileType> WriterSupported
	{
		get
		{
			yield return ExcelXml;
			yield return ExcelXmlMacroEnabled;
#if NETCOREAPP1_0_OR_GREATER
			// not supported on .NET framework versions.
			yield return ExcelBinary;
#endif
		}
	}


	/// <summary>
	/// Gets the ExcelFileType for a given filename.
	/// </summary>
	/// <returns>An ExcelFileType or null.</returns>
	public static ExcelFileType? FindForFilename(string filename)
	{
		var ext = Path.GetExtension(filename);
		return FindForExtension(ext);
	}

	/// <summary>
	/// Gets the ExcelFileType for a given file extension.
	/// </summary>
	/// <returns>An ExcelFileType or null.</returns>
	public static ExcelFileType? FindForExtension(string extension)
	{
		foreach (var type in All)
		{
			if (StringComparer.OrdinalIgnoreCase.Equals(type.Extension, extension))
			{
				return type;
			}
		}
		return null;
	}

	/// <summary>
	/// Gets the ExcelFileType for a given content type.
	/// </summary>
	/// <returns>An ExcelFileType or null.</returns>
	public static ExcelFileType? FindForContentType(string contentType)
	{
		foreach (var type in All)
		{
			if (StringComparer.OrdinalIgnoreCase.Equals(type.ContentType, contentType))
			{
				return type;
			}
		}
		return null;
	}

	/// <summary>
	/// Gets the ExcelFileType for a ExcelWorkbookType.
	/// </summary>
	public static ExcelFileType GetForWorkbookType(ExcelWorkbookType type)
	{
		switch (type)
		{
			case ExcelWorkbookType.Excel:
				return Excel;
			case ExcelWorkbookType.ExcelXml:
				return ExcelXml;
			case ExcelWorkbookType.ExcelBinary:
				return ExcelBinary;
		}
		throw new NotSupportedException();
	}

	/// <summary>
	/// Determines if the string defines a known Excel ContentType.
	/// </summary>
	public static bool IsExcelContentType(string contentType)
	{
		return FindForContentType(contentType) != null;
	}

	private ExcelFileType(string extension, string contentType, ExcelWorkbookType workbookType)
	{
		this.Extension = extension;
		this.ContentType = contentType;
		this.WorkbookType = workbookType;
	}

	/// <summary>
	/// Gets the file extension.
	/// </summary>
	public string Extension { get; }

	/// <summary>
	/// Gets the Content-Type.
	/// </summary>
	public string ContentType { get; }

	/// <summary>
	/// Gets the ExcelWorkbookType.
	/// </summary>
	public ExcelWorkbookType WorkbookType { get; }
}
