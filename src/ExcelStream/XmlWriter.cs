namespace ExcelStream
{
	using System;
	using System.Collections.Generic;
	using System.IO;
	using System.Text;
	using Ionic.Zip;
	using Ionic.Zlib;

	public class XmlWriter : IDisposable
	{
		public void AppendRow(string[] values)
		{
			if (this.disposed)
				throw new ObjectDisposedException(typeof(XmlWriter).Name);

			if (this.saved)
				throw new InvalidOperationException("Unable to write after saving.");

			this.stringBuilder.Clear();

			this.AddRow(values ?? new string[0], ++this.rowNumber); // 1-based row
			this.FlushRow();
		}
		private void AddRow(string[] values, int row)
		{
			this.stringBuilder.AppendFormat(Formats.RowPrefix, row);

			for (var i = 0; i < values.Length; i++)
				this.AddColum(values[i], i);

			this.stringBuilder.Append(Constants.RowSuffix);
		}
		private void AddColum(string value, int column)
		{
			int sharedStringIndex;
			if (!this.sharedStrings.TryGetValue(value, out sharedStringIndex))
			{
				this.sharedStrings[value] = sharedStringIndex = this.sharedStrings.Count; // 0-based
				this.insertionOrder.Add(value);
			}

			this.stringBuilder.AppendFormat(Formats.Cell, (column + 1).ToColumnNumber(), this.rowNumber, sharedStringIndex);
		}
		private void FlushRow()
		{
			for (var i = 0; i < this.stringBuilder.Length; i++)
				this.worksheetWriter.Write(this.stringBuilder[i]);
		}

		public void Save()
		{
			if (this.disposed)
				throw new ObjectDisposedException(typeof(XmlWriter).Name);

			if (this.saved)
				throw new InvalidOperationException("Unable to save after saving.");

			this.saved = true;

			this.WriteWorksheetFooter();
			this.WriteSharedStrings();
			this.WriteMetadata();

			this.zipStream.TryDispose();
			this.sharedStrings.Clear();
			this.insertionOrder.Clear();
			this.stringBuilder.Clear();
		}
		private StreamWriter WriteWorksheetHeader()
		{
			this.zipStream.PutNextEntry(Paths.Sheet1);
			var outputStream = new IndisposableStream(this.zipStream);
			var writer = new StreamWriter(outputStream, DefaultEncoding);
			writer.Write(Constants.WorksheetHeader);
			return writer;
		}
		private void WriteWorksheetFooter()
		{
			this.worksheetWriter.Write(Constants.WorksheetFooter);
			this.worksheetWriter.TryDispose();
		}
		private void WriteSharedStrings()
		{
			this.zipStream.PutNextEntry(Paths.SharedStrings);
			using (var writer = new StreamWriter(new IndisposableStream(this.zipStream), DefaultEncoding))
			{
				writer.Write(Constants.SharedStringsHeader);
				for (var i = 0; i < this.insertionOrder.Count; i++)
				{
					writer.Write(Constants.SharedStringsItemPrefix);
					writer.Write(this.insertionOrder[i].Sanitize());
					writer.Write(Constants.SharedStringsItemSuffix);
				}
				writer.Write(Constants.SharedStringsFooter);
			}
		}
		private void WriteMetadata()
		{
			this.WriteMetadata(Paths.ContentTypes, Metadata.content_types);
			this.WriteMetadata(Paths.RootRelationship, Metadata.rels);
			this.WriteMetadata(Paths.Workbook, Metadata.workbook);
			this.WriteMetadata(Paths.Styles, Metadata.styles);
			this.WriteMetadata(Paths.Theme, Metadata.theme1);
			this.WriteMetadata(Paths.WorkbookRelationship, Metadata.workbook_xml_rels);
		}
		private void WriteMetadata(string filename, byte[] contents)
		{
			this.zipStream.PutNextEntry(filename);
			this.zipStream.Write(contents, 0, contents.Length);
		}

		public XmlWriter(string outputPath) : this(File.Open(outputPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
		{
		}
		public XmlWriter(Stream outputStream)
		{
			if (outputStream == null)
				throw new ArgumentNullException("outputStream");

			if (!outputStream.CanWrite)
				throw new InvalidOperationException("Unable to write to stream provided.");

			try
			{
				this.sharedStrings = new Dictionary<string, int>();
				this.insertionOrder = new List<string>();

				this.zipStream = new ZipOutputStream(outputStream, true)
				{
					CompressionLevel = CompressionLevel.Level9,
					CompressionMethod = CompressionMethod.Deflate,
				};

				this.worksheetWriter = this.WriteWorksheetHeader();
			}
			catch
			{
				this.Dispose();
				throw;
			}
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

			this.zipStream.TryDispose();
			this.worksheetWriter.TryDispose();

			this.stringBuilder.Clear();
			this.sharedStrings.Clear();
			this.insertionOrder.Clear();
		}

		private static readonly Encoding DefaultEncoding = new UTF8Encoding(true);
		private readonly StringBuilder stringBuilder = new StringBuilder();
		private readonly ZipOutputStream zipStream;
		private readonly Dictionary<string, int> sharedStrings;
		private readonly List<string> insertionOrder;
		private readonly StreamWriter worksheetWriter;
		private bool saved;
		private bool disposed;
		private int rowNumber;
	}
}