namespace ExcelStream
{
	using System;

	internal static class Constants
	{
		public const string WorksheetPrefix = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"><sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"/></sheetViews><sheetData>";
		public const string WorksheetSuffix = "</sheetData></worksheet>";
		public const string SharedStringsPrefix = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">";
		public const string SharedStringsSuffix = "</sst>";
	}

	internal static class Formats
	{
		public const string Row = "<row r=\"{0}\">{1}</row>"; // 0: row number, 1: cell contents
		public const string Cell = "<c r=\"{0}\" s=\"1\" t=\"s\"><v>{1}</v></c>"; // 0: A1, B1, C1, etc. 1: shared string integer
		public const string SharedString = "<si><t>{0}</t></si>";
	}

	internal static class MimeTypes
	{
		public const string Worksheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
		public const string SharedStrings = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
		public const string Workbook = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
		public const string Theme = "application/vnd.openxmlformats-officedocument.theme+xml";
		public const string Style = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
	}

	internal static class Uris
	{
		public static readonly Uri Sheet11 = new Uri("/xl/worksheets/sheet1.xml", UriKind.Relative);
		public static readonly Uri SharedStrings = new Uri("/xl/sharedStrings.xml", UriKind.Relative);
		public static readonly Uri Workbook = new Uri("/xl/workbook.xml", UriKind.Relative);
		public static readonly Uri Theme = new Uri("/xl/theme/theme1.xml", UriKind.Relative);
		public static readonly Uri Styles = new Uri("/xl/styles.xml", UriKind.Relative);
	}
}