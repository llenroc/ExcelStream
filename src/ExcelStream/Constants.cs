namespace ExcelStream
{
	internal static class Constants
	{
		public const string WorksheetHeader = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"><sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews><sheetData>";
		public const string WorksheetFooter = "</sheetData></worksheet>";
		public const string SharedStringsHeader = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">";
		public const string SharedStringsFooter = "</sst>";

		public const string RowSuffix = "</row>";
		public const string SharedStringsItemPrefix = "<si><t>";
		public const string SharedStringsItemSuffix = "</t></si>";
	}
	internal static class Formats
	{
		public const string RowPrefix = "<row r=\"{0}\">"; // 0: row number
		public const string Cell = "<c r=\"{0}{1}\" s=\"1\" t=\"s\"><v>{2}</v></c>"; // 0: column, 1: row, 2: shared string integer
	}
	internal static class Paths
	{
		public const string Sheet1 = "xl/worksheets/sheet1.xml";
		public const string SharedStrings = "xl/sharedStrings.xml";
	}
}