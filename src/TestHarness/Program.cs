namespace TestHarness
{
	using System;
	using System.Globalization;
	using ExcelStream;

	class Program
	{
		public static void Main()
		{
			using (var writer = new ExcelWriter("c:/" + DateTime.UtcNow.Ticks + ".xlsx"))
			{
				var cells = new string[11];

				writer.AppendRow(new[] { "Hello", "World!" });
				for (var i = 0; i < 101; i++)
				{
					for (var x = 0; x < cells.Length; x++)
						cells[x] = (i * x).ToString(CultureInfo.InvariantCulture);

					writer.AppendRow(cells);
				}

				writer.Save();
			}
		}
	}
}