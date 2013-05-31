namespace TestHarness
{
	using System;
	using ExcelStream;

	class Program
	{
		public static void Main()
		{
			var streamer = new ExcelWriter("c:/" + Guid.NewGuid() + ".zip");
			streamer.Write(new[] { "Hello", "World" });
			streamer.Write(new[] { "second", "row" });
			streamer.Save();
			streamer.Dispose();
		}
	}
}