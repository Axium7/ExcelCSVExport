namespace ExcelCSVExport.Services;

public class SerAppSession
{
	private IWebHostEnvironment WebHostEnvironment;

	public SerAppSession(IWebHostEnvironment webHostEnvironment)
	{
		WebHostEnvironment = webHostEnvironment ?? throw new ArgumentNullException(nameof(webHostEnvironment));
	}

	public string StrAppTitle { get; private set; } = "ExcelCSVExport";

	public string WebRootExportFilesPath
	{
		get
		{
			var path = Path.Combine(WebHostEnvironment.WebRootPath, "exportfiles");
			if (!Directory.Exists(path))
			{
				Directory.CreateDirectory(path);
			}
			return path;
		}
	}

}

