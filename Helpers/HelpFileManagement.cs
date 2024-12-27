using System.Diagnostics;
using ExcelCSVExport.Enums;

namespace ExcelCSVExport.Helpers;

public static class HelpFileManagement
{
	public static string CreateFileName(params string[] parts)
	{
		return string.Join(" ", parts.Where(s => !string.IsNullOrEmpty(s)));
	}

	public static string CreateFullFileName(string strFileName, ExportFormat format)
	{
		switch (format)
		{
			case ExportFormat.Excel:
				strFileName += ".xlsx";
				break;
			case ExportFormat.CSV:
				strFileName += ".csv";
				break;
		}

		return strFileName;
	}
}