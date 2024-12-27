using System.Diagnostics;
using ExcelCSVExport.Enums;

namespace ExcelCSVExport.Helpers;

public static class HelpFileManagement
{
	public static string DownloadsFullPath(string strFileName)
	{
		// Local Downloads Folder
		return Path.Combine(
			Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
			"Downloads",
			strFileName
		);
	}

	public static string DesktopFullPath(string strFileName)
	{
		// Desktop
		return Path.Combine(
			Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
			"Desktop",
			strFileName
		);
	}

	public static async Task OpenFileAsync(string strFileFullPath)
	{
		// Open the file with the default application asynchronously
		await Task.Run(() => Process.Start(new ProcessStartInfo
		{
			FileName = strFileFullPath,
			UseShellExecute = true
		}));
	}

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
			case ExportFormat.PDF:
				strFileName += ".pdf";
				break;
			default:
				throw new ArgumentOutOfRangeException(nameof(format), $"Unsupported format: {format}");
		}

		return strFileName;
	}

	public static async Task<bool> IsFileOpenAsync(string filePath)
	{
		try
		{
			using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
			{
				// The file can be accessed, so it is not open
				return false;
			}
		}
		catch (IOException)
		{
			// The file is in use if an IOException is thrown
			return true;
		}
	}
}