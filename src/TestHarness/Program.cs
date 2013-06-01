namespace TestHarness
{
	using System;
	using ExcelStream;

	class Program
	{
		public static void Main()
		{
			var writer = new ExcelWriter("c:/" + DateTime.UtcNow.Ticks + ".xlsx");
			var cells = new string[11];

			writer.AppendRow(new string[] { "<Hello>", "\"'&World!" });
			for (var i = 0; i < 101; i++)
			{
				for (var x = 0; x < cells.Length; x++)
					cells[x] = (i * x).ToString();

				writer.AppendRow(cells);
			}
				
			writer.Save();
			writer.Dispose();
		}
	}
}