namespace ExcelStream
{
	using System;
	using System.Collections.Generic;
	using System.IO;
	using System.Text;
	using Ionic.Zip;
	using Ionic.Zlib;

	public sealed class ExcelWriter : IExcelWriter
	{
		public void AppendRow(string[] values)
		{
			if (this.disposed)
				throw new ObjectDisposedException(typeof(ExcelWriter).Name);

			if (this.saved)
				throw new InvalidOperationException("Unable to write after saving.");

			this.stringBuilder.Clear();

			this.AppendRow(values ?? new string[0], ++this.maxRow); // 1-based row
			this.FlushRow();
		}
		private void AppendRow(string[] values, int row)
		{
			this.stringBuilder.AppendFormat(Formats.RowPrefix, row);

			for (var i = 0; i < values.Length; i++)
				this.AppendColumn(values[i], i);

			this.stringBuilder.Append(Constants.RowSuffix);
		}
		private void AppendColumn(string value, int column)
		{
			int sharedStringIndex;
			if (!this.sharedStrings.TryGetValue(value, out sharedStringIndex))
			{
				this.sharedStrings[value] = sharedStringIndex = this.sharedStrings.Count; // 0-based
				this.insertionOrder.Add(value);
			}

			this.stringBuilder.AppendFormat(Formats.Cell, (column + 1).ToColumnNumber(), this.maxRow, sharedStringIndex);
		}
		private void FlushRow()
		{
			for (var i = 0; i < this.stringBuilder.Length; i++)
				this.worksheetWriter.Write(this.stringBuilder[i]);
		}

		public void Save()
		{
			if (this.disposed)
				throw new ObjectDisposedException(typeof(ExcelWriter).Name);

			if (this.saved)
				throw new InvalidOperationException("Unable to save after saving.");

			this.saved = true;

			this.WriteWorksheetFooter();
			this.WriteSharedStrings();
			this.WriteMetadata();

			this.Close();
		}
		private StreamWriter WriteWorksheetHeader()
		{
			this.zipStream.PutNextEntry(Paths.Sheet1);
			var writer = new StreamWriter(new IndisposableStream(this.zipStream), DefaultEncoding);
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
			using (var source = new MemoryStream(Metadata.metadata))
			using (var unzipped = new ZipInputStream(source))
			{
				ZipEntry entry;
				var buffer = new byte[1024 * 32]; // this is larger than the metadata file

				while ((entry = unzipped.GetNextEntry()) != null)
				{
					if (entry.UncompressedSize == 0)
						continue;

					var size = (int)entry.UncompressedSize;
					unzipped.Read(buffer, 0, size);
					this.zipStream.PutNextEntry(entry.FileName);
					this.zipStream.Write(buffer, 0, size);
				}
			}
		}
		private void Close()
		{
			this.zipStream.TryDispose();
			this.worksheetWriter.TryDispose();

			if (this.disposeOutputStream)
				this.outputStream.TryDispose();

			this.stringBuilder.Clear();
			this.sharedStrings.Clear();
			this.insertionOrder.Clear();
		}

		public ExcelWriter(string outputPath) : this(File.Open(outputPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None), true)
		{
		}
		public ExcelWriter(Stream outputStream, bool dispose = false)
		{
			if (outputStream == null)
				throw new ArgumentNullException("outputStream");

			if (!outputStream.CanWrite)
				throw new InvalidOperationException("Unable to write to the stream provided.");

			try
			{
				this.disposeOutputStream = dispose;
				this.outputStream = outputStream;
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
		~ExcelWriter()
		{
			this.Dispose(false);
		}

		public void Dispose()
		{
			this.Dispose(true);
			GC.SuppressFinalize(this);
		}
		private void Dispose(bool disposing)
		{
			if (disposing || this.disposed)
				return;

			this.disposed = true;
			this.Close();
		}

		private static readonly Encoding DefaultEncoding = new UTF8Encoding(true);
		private readonly StringBuilder stringBuilder = new StringBuilder();
		private readonly Dictionary<string, int> sharedStrings = new Dictionary<string, int>(1024 * 1024);
		private readonly List<string> insertionOrder = new List<string>(1024 * 1024);
		private readonly bool disposeOutputStream;
		private readonly Stream outputStream;
		private readonly ZipOutputStream zipStream;
		private readonly StreamWriter worksheetWriter;
		private bool saved;
		private bool disposed;
		private int maxRow;
	}
}