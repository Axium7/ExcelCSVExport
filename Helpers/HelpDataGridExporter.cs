using MudBlazor;
using ExcelCSVExport.HelpersExternal;

namespace ExcelCSVExport.Helpers;

public static class HelpDataGridExporter<T>
{
	public static async Task ExportExcelAsync(MudDataGrid<T> mudGrid, string strFileFullPath)
	{
		// Retrieve the data from the grid
		var data = await ExtDataGridExporter.GetTableDataAsync(
			mudGrid.RenderedColumns,
			mudGrid.FilteredItems
		);
		var fileContentAsByteArray = await ExtDataGridExporter.GenerateExcelAsync(data);

		// Save the file asynchronously
		await SaveExportToFileAsync(fileContentAsByteArray, strFileFullPath);
		
		//Format the Excel file as necessary
	}

	public static async Task ExportCSVAsync(MudDataGrid<T> mudGrid, string strFileFullPath)
	{
		// Retrieve the data from the grid
		var data = await ExtDataGridExporter.GetTableDataAsync(
			mudGrid.RenderedColumns,
			mudGrid.FilteredItems
		);

		var fileContentAsByteArray = await ExtDataGridExporter.GenerateCSVAsync(data);

		// Save the file asynchronously
		await SaveExportToFileAsync(fileContentAsByteArray, strFileFullPath);

	}
}
