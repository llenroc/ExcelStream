namespace ExcelStream
{
	using System;
	using System.IO;
	using System.Security;
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

			this.rowSequence++;
			this.WriteRow(values ?? new string[0]); // 1-based row
		}
		private void WriteRow(string[] values)
		{
			this.worksheetWriter.Write(Constants.RowPrefix);

			for (var i = 0; i < values.Length; i++)
				this.WriteColumn(values[i], i + 1);

			this.worksheetWriter.Write(Constants.RowSuffix);
		}
		private void WriteColumn(string value, int cellSequence)
		{
			this.worksheetWriter.Write(Formats.CellPrefix, cellSequence.ToColumnNumber(), this.rowSequence);
			this.worksheetWriter.Write(SecurityElement.Escape(value ?? string.Empty));
			this.worksheetWriter.Write(Constants.CellSuffix);
		}

		public void Save()
		{
			if (this.disposed)
				throw new ObjectDisposedException(typeof(ExcelWriter).Name);

			if (this.saved)
				throw new InvalidOperationException("Unable to save after saving.");

			this.saved = true;

			this.WriteWorksheetFooter();
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

			if (!this.leaveOutputStreamOpen)
				this.outputStream.TryDispose();
		}

		public ExcelWriter(string outputPath) : this(File.Open(outputPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None), false)
		{
		}
		public ExcelWriter(Stream outputStream, bool leaveOpen = true)
		{
			if (outputStream == null)
				throw new ArgumentNullException("outputStream");

			if (!outputStream.CanWrite)
				throw new InvalidOperationException("Unable to write to the stream provided.");

			try
			{
				this.leaveOutputStreamOpen = leaveOpen;
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
		private readonly bool leaveOutputStreamOpen;
		private readonly Stream outputStream;
		private readonly ZipOutputStream zipStream;
		private readonly StreamWriter worksheetWriter;
		private bool saved;
		private bool disposed;
		private int rowSequence;
	}
}
