namespace ExcelStream
{
	using System;
	using System.Collections.Generic;
	using System.IO;
	using System.IO.Packaging;
	using System.Text;

	public class ExcelWriter : IDisposable
	{
		public void Write(string[] values)
		{
			if (this.disposed)
				throw new ObjectDisposedException(typeof(ExcelWriter).Name);

			if (this.saved)
				throw new InvalidOperationException("Unable to write after saving.");

			this.rowBuilder.Clear();
			this.cellBuilder.Clear();

			this.rowNumber++; // 1-based

			values = values ?? new string[0];
			for (var i = 0; i < values.Length; i++)
			{
				this.columnNumber = Math.Max(this.columnNumber, i + 1); // 1-based
				var id = AsColumnNumber(i) + this.rowBuilder;
				var contents = Sanitize(values[i]);
				this.cellBuilder.AppendFormat(Constants.CellFormat, id, contents);
			}

			this.rowBuilder.AppendFormat(Constants.RowFormat, this.rowNumber, this.cellBuilder);
			for (var i = 0; i < this.rowBuilder.Length; i++)
				this.worksheetWriter.Write(this.rowBuilder[i]);
		}
		public void Save()
		{
			if (this.disposed)
				throw new ObjectDisposedException(typeof(ExcelWriter).Name);

			if (this.saved)
				throw new InvalidOperationException("Unable to save after saving.");

			this.saved = true;

			this.worksheetWriter.Flush();
			// TOOD: write shared strings
		}

		private static string AsColumnNumber(int columnNumber)
		{
			var dividend = columnNumber;
			var columnName = string.Empty;

			while (dividend > 0)
			{
				var modulo = (dividend - 1) % 26;
				columnName = Convert.ToChar(65 + modulo) + columnName;
				dividend = (dividend - modulo) / 26;
			}

			return columnName;
		}
		private static string Sanitize(string value)
		{
			if (string.IsNullOrEmpty(value))
				return string.Empty;

			return value; // TODO
		}

		public ExcelWriter(string outputPath, int rows = 0) : this(File.Open(outputPath, FileMode.Create, FileAccess.Write, FileShare.None), rows)
		{
		}
		public ExcelWriter(Stream outputStream, int rows = 0)
		{
			if (outputStream == null)
				throw new ArgumentNullException("outputStream");

			if (!outputStream.CanWrite)
				throw new InvalidOperationException("Unable to write to stream provided.");

			if (rows < 0)
				throw new ArgumentOutOfRangeException("rows", "Capacity must be positive.");

			this.sharedStrings = rows > 0 ? new Dictionary<string, int>(rows) : new Dictionary<string, int>();
			this.outputPackage = Package.Open(outputStream);
			this.worksheet = this.outputPackage.CreatePart(SheetUri, MimeType, CompressionOption.Maximum);
			this.worksheetWriter = new StreamWriter(this.worksheet.GetStream(), Encoding.UTF8);

			// TODO: write all of the "filler" (static) files to the stream
		}

		public void Dispose()
		{
			this.Dispose(true);
			GC.SuppressFinalize(this);
		}
		protected virtual void Dispose(bool disposing)
		{
			if (disposing || this.disposed)
				return;

			this.disposed = true;
			this.outputPackage.Flush();
			this.cellBuilder.Clear();
			this.rowBuilder.Clear();
		}

		private const string MimeType = System.Net.Mime.MediaTypeNames.Application.Octet;
		private static readonly Uri SheetUri = new Uri("xl/worksheets/sheet1.xml", UriKind.Relative);
		private readonly StringBuilder rowBuilder = new StringBuilder();
		private readonly StringBuilder cellBuilder = new StringBuilder();
		private readonly Dictionary<string, int> sharedStrings;
		private readonly Package outputPackage;
		private readonly PackagePart worksheet;
		private readonly StreamWriter worksheetWriter;
		private bool saved;
		private bool disposed;
		private int columnNumber;
		private int rowNumber;
	}
}