﻿namespace TestHarness
{
	using System;
	using ExcelStream;

	class Program
	{
		public static void Main()
		{
			var writer = new XmlWriter("c:/" + Guid.NewGuid() + ".xlsx");
			var cells = new string[11];

			for (var i = 0; i < 10001; i++)
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