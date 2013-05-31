namespace ExcelStream
{
	internal static class Constants
	{
		public static string RowFormat = "<row r=\"{0}\">{1}</row>"; // 0: row number, 1: cell contents
		public static string CellFormat = "<c r=\"{0}\" s=\"1\" t=\"s\"><v>{1}</v></c>"; // 0: A1, B1, C1, etc. 1: shared string integer
	}
}