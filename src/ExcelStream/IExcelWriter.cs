namespace ExcelStream
{
	using System;

	public interface IExcelWriter : IDisposable
	{
		void AppendRow(string[] values);
		void Save();
	}
}