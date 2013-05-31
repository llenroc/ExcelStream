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
				var id = (i + 1).ToColumnNumber() + this.rowNumber;

				var sanitized = values[i].Sanitize();

				int sharedStringIndex;
				if (!this.sharedStrings.TryGetValue(sanitized, out sharedStringIndex))
				{
					this.sharedStrings[sanitized] = sharedStringIndex = this.sharedStrings.Count; // 0-based
					this.sharedStringsWriter.Write(Formats.SharedString, sanitized);
				}

				this.cellBuilder.AppendFormat(Formats.Cell, id, sharedStringIndex);
			}

			this.rowBuilder.AppendFormat(Formats.Row, this.rowNumber, this.cellBuilder);
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

			this.worksheetWriter.Write(Constants.WorksheetSuffix);
			this.sharedStringsWriter.Write(Constants.SharedStringsSuffix);

			this.worksheetWriter.Dispose();
			this.sharedStringsWriter.Dispose();

			this.outputPackage.Close();
			this.sharedStrings.Clear();
		}

		public ExcelWriter(string outputPath, int rows = 0) : this(File.Open(outputPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None), rows)
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
			this.outputPackage = Package.Open(outputStream, FileMode.Create);

			this.worksheetPart = this.outputPackage.CreatePart(Uris.Sheet1, MimeTypes.Worksheet, CompressionOption.Maximum);
			this.sharedStringsPart = this.outputPackage.CreatePart(Uris.SharedStrings, MimeTypes.SharedStrings, CompressionOption.Maximum);

			this.worksheetWriter = new StreamWriter(this.worksheetPart.GetStream(), Encoding.UTF8);
			this.sharedStringsWriter = new StreamWriter(this.sharedStringsPart.GetStream(), Encoding.UTF8);

			this.worksheetWriter.Write(Constants.WorksheetPrefix);
			this.sharedStringsWriter.Write(Constants.SharedStringsPrefix);

			this.WritePreamble();
		}
		private void WritePreamble()
		{
			this.WritePart(Uris.Styles, FileResources.styles, MimeTypes.Style);
			this.WritePart(Uris.Theme, FileResources.theme1, MimeTypes.Theme);
			this.WritePart(Uris.Workbook, FileResources.workbook, MimeTypes.Workbook);

			this.WritePart(Uris.RootRelationship, FileResources.rels, "application/octet-stream");
			this.WritePart(Uris.WorkbookRelationship, FileResources.workbook_xml_rels, "application/octet-stream");
		}
		private void WritePart(Uri location, string content, string mimeType)
		{
			var part = this.outputPackage.CreatePart(location, mimeType, CompressionOption.Maximum);
			using (var writer = new StreamWriter(part.GetStream(), Encoding.UTF8))
				writer.Write(content);
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
			this.sharedStrings.Clear();
		}
		
		private readonly StringBuilder rowBuilder = new StringBuilder();
		private readonly StringBuilder cellBuilder = new StringBuilder();
		private readonly Dictionary<string, int> sharedStrings;
		private readonly Package outputPackage;
		private readonly PackagePart sharedStringsPart;
		private readonly PackagePart worksheetPart;
		private readonly StreamWriter sharedStringsWriter;
		private readonly StreamWriter worksheetWriter;
		private bool saved;
		private bool disposed;
		private int columnNumber;
		private int rowNumber;
	}
}