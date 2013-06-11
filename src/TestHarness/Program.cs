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
				var cells = new string[25];

				writer.AppendRow(new[] { "Hello", "World!" });
				for (var i = 0; i < 1000001; i++)
				{
					if (i > 0 && i % 10000 == 0)
					{
						Console.Clear();
						Console.WriteLine(i);
					}

					for (var x = 0; x < cells.Length; x++)
						cells[x] = (i * x).ToString(CultureInfo.InvariantCulture);

					writer.AppendRow(cells);
				}

				writer.Save();
			}
		}
	}
}