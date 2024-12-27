//Credits: https://gist.github.com/Apflkuacha/3eaa55ca52675329ce76f5cd725e472e#gistcomment-5358988

using System.Globalization;
using System.Reflection;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using MudBlazor;
using Color = DocumentFormat.OpenXml.Spreadsheet.Color;

namespace ExcelCSVExport.HelpersExternal;

//Credits: https://gist.github.com/Apflkuacha/3eaa55ca52675329ce76f5cd725e472e#gistcomment-5358988

public static class ExtDataGridExporter
{
	public static async Task<TableData> GetTableDataAsync<T>(
		IEnumerable<Column<T>> columns,
		IEnumerable<T> items
	)
	{
		var tableData = new TableData { SheetName = "Items" };

		//Generate the Header Cells.
		var header = new List<Cell>();
		foreach (var column in columns)
		{
			if (!column.Hidden && !string.IsNullOrEmpty(column.PropertyName))
			{
				header.Add(
					new Cell()
					{
						CellValue = new CellValue(column.Title),
						DataType = CellValues.String,
					}
				);
			}
		}

		tableData.Cells.Add(header);

		var type = items.FirstOrDefault()?.GetType();
		foreach (var item in items)
		{
			if (item == null || type == null)
			{
				continue;
			}

			List<Cell> row = new List<Cell>();
			foreach (var column in columns)
			{
				if (!column.Hidden && !string.IsNullOrEmpty(column.PropertyName))
				{
					row.Add(GetCell(item, type.GetProperty(column.PropertyName)));
				}
			}
			tableData.Cells.Add(row);
		}
		return await Task.FromResult(tableData);
	}

	private static Cell GetCell(object? item, PropertyInfo? prop)
	{
		var cell = new Cell();
		if (item == null || prop == null)
		{
			return cell;
		}

		var value = prop.GetValue(item);
		var valueType = prop.PropertyType;
		var stringValue = value?.ToString()?.Trim() ?? "";

		var underlyingType =
			valueType.IsGenericType && valueType.GetGenericTypeDefinition() == typeof(Nullable<>)
				? Nullable.GetUnderlyingType(valueType)
				: valueType;

		var typeCode = Type.GetTypeCode(underlyingType);

		if (typeCode == TypeCode.DateTime)
		{
			if (!string.IsNullOrWhiteSpace(stringValue) && value != null)
			{
				cell.CellValue = new CellValue()
				{
					Text = ((DateTime)value).ToOADate().ToString(CultureInfo.InvariantCulture),
				};
				cell.DataType = CellValues.Number;
				cell.StyleIndex = (UInt32Value)1U;
			}
		}
		else if (typeCode == TypeCode.Boolean)
		{
			cell.CellValue = new CellValue(stringValue.ToLower());
			cell.DataType = CellValues.Boolean;
		}
		else if (IsNumeric(typeCode) && underlyingType?.IsEnum != true)
		{
			if (value is double doubleValue)
			{
				stringValue = Math.Round(doubleValue, 2).ToString().Replace(',', '.');
			}
			else if (value != null)
			{
				stringValue = Convert.ToString(value, CultureInfo.InvariantCulture);
			}

			cell.CellValue = new CellValue(stringValue ?? "");
			cell.DataType = CellValues.Number;
		}
		else
		{
			cell.CellValue = new CellValue(stringValue);
			cell.DataType = CellValues.String;
		}

		return cell;
	}

	public static async Task<byte[]> GenerateCSVAsync(TableData data)
	{
		var sb = new StringBuilder();
		foreach (var rowData in data.Cells)
		{
			var row = new List<string>();
			foreach (var cell in rowData)
			{
				var value = cell.CellValue?.Text?.Trim() ?? "";
				if (
					cell.StyleIndex?.Value == 1
					&& cell.CellValue?.TryGetDouble(out var number) == true
				)
				{
					var date1 = DateTime.FromOADate(number);
					value = date1.ToString("dd.MM.yyyy HH:mm");
				}
				if (
					cell.DataType != null
					&& cell.DataType == CellValues.Date
					&& cell.CellValue?.TryGetDateTime(out var date) == true
				)
				{
					value = date.ToString("dd.MM.yyyy HH:mm");
				}
				if (value.Contains(','))
				{
					value = $"\"{value}\"";
				}

				row.Add(value);
			}
			sb.AppendLine(string.Join(",", row.ToArray()));
		}
		return await Task.FromResult(Encoding.UTF8.GetBytes(sb.ToString()));
	}

	public static async Task<byte[]> GenerateExcelAsync(TableData data)
	{
		using var stream = new MemoryStream();
		using var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
		{
			var workbookPart = document.AddWorkbookPart();
			workbookPart.Workbook = new Workbook();
			var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

			var sheetData = new SheetData();
			worksheetPart.Worksheet = new Worksheet(sheetData);

			var workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
			GenerateWorkbookStylesPartContent(workbookStylesPart);
			if (document.WorkbookPart == null)
			{
				return stream.ToArray();
			}

			var sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());

			var sheet = new Sheet()
			{
				Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
				SheetId = 1,
				Name = data.SheetName ?? "Sheet 1",
			};
			sheets.AppendChild(sheet);

			AppendDataToSheet(sheetData, data);

			workbookPart.Workbook.Save();
			document.Dispose();
		}
		return await Task.FromResult(stream.ToArray()); // Simulate async work
	}

	private static void AppendDataToSheet(SheetData sheetData, TableData data)
	{
		foreach (var rowData in data.Cells)
		{
			var row = new DocumentFormat.OpenXml.Spreadsheet.Row();
			sheetData.AppendChild(row);
			foreach (var cell in rowData)
			{
				row.AppendChild(cell);
			}
		}
	}

	private static bool IsNumeric(TypeCode typeCode)
	{
		return typeCode switch
		{
			TypeCode.Decimal
			or TypeCode.Double
			or TypeCode.Int16
			or TypeCode.Int32
			or TypeCode.Int64
			or TypeCode.UInt16
			or TypeCode.UInt32
			or TypeCode.UInt64 => true,
			_ => false,
		};
	}

	private static void GenerateWorkbookStylesPartContent(WorkbookStylesPart workbookStylesPart1)
	{
		var stylesheet1 = new Stylesheet()
		{
			MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac x16r2 xr" },
		};
		stylesheet1.AddNamespaceDeclaration(
			"mc",
			"http://schemas.openxmlformats.org/markup-compatibility/2006"
		);
		stylesheet1.AddNamespaceDeclaration(
			"x14ac",
			"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
		);
		stylesheet1.AddNamespaceDeclaration(
			"x16r2",
			"http://schemas.microsoft.com/office/spreadsheetml/2015/02/main"
		);
		stylesheet1.AddNamespaceDeclaration(
			"xr",
			"http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
		);

		var fonts1 = new Fonts() { Count = (UInt32Value)1U, KnownFonts = true };

		var font1 = new Font();
		var fontSize1 = new FontSize() { Val = 11D };
		var color1 = new Color() { Theme = (UInt32Value)1U };
		var fontName1 = new FontName() { Val = "Calibri" };
		var fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
		var fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

		font1.Append(fontSize1);
		font1.Append(color1);
		font1.Append(fontName1);
		font1.Append(fontFamilyNumbering1);
		font1.Append(fontScheme1);
		fonts1.Append(font1);

		var fills1 = new Fills() { Count = (UInt32Value)2U };

		var fill1 = new Fill();
		var patternFill1 = new PatternFill() { PatternType = PatternValues.None };

		fill1.Append(patternFill1);

		var fill2 = new Fill();
		var patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

		fill2.Append(patternFill2);

		fills1.Append(fill1);
		fills1.Append(fill2);

		var borders1 = new Borders() { Count = (UInt32Value)1U };

		var border1 = new Border();
		var leftBorder1 = new LeftBorder();
		var rightBorder1 = new RightBorder();
		var topBorder1 = new TopBorder();
		var bottomBorder1 = new BottomBorder();
		var diagonalBorder1 = new DiagonalBorder();

		border1.Append(leftBorder1);
		border1.Append(rightBorder1);
		border1.Append(topBorder1);
		border1.Append(bottomBorder1);
		border1.Append(diagonalBorder1);

		borders1.Append(border1);

		var cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
		var cellFormat1 = new CellFormat()
		{
			NumberFormatId = (UInt32Value)0U,
			FontId = (UInt32Value)0U,
			FillId = (UInt32Value)0U,
			BorderId = (UInt32Value)0U,
		};

		cellStyleFormats1.Append(cellFormat1);

		var cellFormats1 = new CellFormats() { Count = (UInt32Value)2U };
		var cellFormat2 = new CellFormat()
		{
			NumberFormatId = (UInt32Value)0U,
			FontId = (UInt32Value)0U,
			FillId = (UInt32Value)0U,
			BorderId = (UInt32Value)0U,
			FormatId = (UInt32Value)0U,
		};
		var cellFormat3 = new CellFormat()
		{
			NumberFormatId = (UInt32Value)14U,
			FontId = (UInt32Value)0U,
			FillId = (UInt32Value)0U,
			BorderId = (UInt32Value)0U,
			FormatId = (UInt32Value)0U,
			ApplyNumberFormat = true,
		};

		cellFormats1.Append(cellFormat2);
		cellFormats1.Append(cellFormat3);

		var cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
		var cellStyle1 = new CellStyle()
		{
			Name = "Normal",
			FormatId = (UInt32Value)0U,
			BuiltinId = (UInt32Value)0U,
		};

		cellStyles1.Append(cellStyle1);
		var differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
		var tableStyles1 = new TableStyles()
		{
			Count = (UInt32Value)0U,
			DefaultTableStyle = "TableStyleMedium2",
			DefaultPivotStyle = "PivotStyleLight16",
		};

		var stylesheetExtensionList1 = new StylesheetExtensionList();

		var stylesheetExtension1 = new StylesheetExtension()
		{
			Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}",
		};
		stylesheetExtension1.AddNamespaceDeclaration(
			"x14",
			"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
		);

		var stylesheetExtension2 = new StylesheetExtension()
		{
			Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}",
		};
		stylesheetExtension2.AddNamespaceDeclaration(
			"x15",
			"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"
		);

		OpenXmlUnknownElement openXmlUnknownElement4 = workbookStylesPart1.CreateUnknownElement(
			"<x15:timelineStyles defaultTimelineStyle=\"TimeSlicerStyleLight1\" xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\" />"
		);

		stylesheetExtension2.Append(openXmlUnknownElement4);

		stylesheetExtensionList1.Append(stylesheetExtension1);
		stylesheetExtensionList1.Append(stylesheetExtension2);

		stylesheet1.Append(fonts1);
		stylesheet1.Append(fills1);
		stylesheet1.Append(borders1);
		stylesheet1.Append(cellStyleFormats1);
		stylesheet1.Append(cellFormats1);
		stylesheet1.Append(cellStyles1);
		stylesheet1.Append(differentialFormats1);
		stylesheet1.Append(tableStyles1);
		stylesheet1.Append(stylesheetExtensionList1);

		workbookStylesPart1.Stylesheet = stylesheet1;
	}

	public class TableData
	{
		public List<List<Cell>> Cells { get; set; } = [];
		public string SheetName { get; set; } = "";
	}
}
