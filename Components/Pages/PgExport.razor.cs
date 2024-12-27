using ExcelCSVExport.Helpers;
using ExcelCSVExport.Enums;
using Microsoft.JSInterop;
using MudBlazor;

namespace ExcelCSVExport.Components.Pages;
public partial class PgExport
{
	private List<Element> lstElements;
	private MudDataGrid<Element> refMudDataGrid;

	protected override async Task OnInitializedAsync()
	{
		lstElements = ElementsList.GetElements();
	}

	private async Task ExportDataAsync(ExportFormat format)
	{
		#region Validation

		//Assign a file name
		string strFileName = HelpFileManagement.CreateFileName("MyFileName");

		//Add Suffix
		strFileName = HelpFileManagement.CreateFullFileName(strFileName, format);

		string strFullFilePath = Path.Combine(serApp.WebRootExportFilesPath, strFileName);

		switch (format)
		{
			case ExportFormat.Excel:
				await HelpDataGridExporter<Element>.ExportExcelAsync(refMudDataGrid, strFullFilePath);
				break;
			case ExportFormat.CSV:
				await HelpDataGridExporter<Element>.ExportCSVAsync(refMudDataGrid, strFullFilePath);
				break;
		}

		// Generate the URL for the file
		string strFileUrl = $"/exportfiles/{strFileName}";

		// Call the JavaScript function to download the file
		await JSRuntime.InvokeVoidAsync("downLoadFile", strFileName, strFileUrl);
	}
}