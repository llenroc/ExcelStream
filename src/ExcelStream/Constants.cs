namespace ExcelStream
{
	internal static class Constants
	{
		public const string WorksheetHeader = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"><sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews><sheetData>";
		public const string WorksheetFooter = "</sheetData></worksheet>";

		public const string RowPrefix = "<row>";
		public const string RowSuffix = "</row>";

		public const string CellSuffix = "</t></is></c>";
	}
	internal static class Formats
	{
		public const string CellPrefix = "<c r=\"{0}{1}\" t=\"inlineStr\"><is><t>"; // 0: column, 1: row
	}
	internal static class Paths
	{
		public const string Sheet1 = "xl/worksheets/sheet1.xml";
	}
}