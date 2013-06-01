namespace ExcelStream
{
	using System;
	using System.IO;
	using System.Runtime.Remoting;
	using System.Security;

	internal static class ExtensionMethods
	{
		public static string Sanitize(this string value)
		{
			return SecurityElement.Escape(value ?? string.Empty);
		}
		public static string ToColumnNumber(this int columnNumber)
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
		public static IDisposable TryDispose(this IDisposable resource)
		{
			try
			{
				if (resource != null)
					resource.Dispose();

				return null;
			}
			catch
			{
				return resource;
			}
		}
	}

	internal sealed class IndisposableStream : Stream
	{
		public Stream BaseStream
		{
			get { return this.stream; }
		}

		public override IAsyncResult BeginRead(byte[] buffer, int offset, int count, AsyncCallback callback, object state)
		{
			return this.stream.BeginRead(buffer, offset, count, callback, state);
		}
		public override IAsyncResult BeginWrite(byte[] buffer, int offset, int count, AsyncCallback callback, object state)
		{
			return this.stream.BeginWrite(buffer, offset, count, callback, state);
		}
		public override bool CanRead
		{
			get { return this.stream.CanRead; }
		}
		public override bool CanSeek
		{
			get { return this.stream.CanSeek; }
		}
		public override bool CanWrite
		{
			get { return this.stream.CanWrite; }
		}
		public override void Close()
		{
			this.stream.Flush();
		}
		public override ObjRef CreateObjRef(Type requestedType)
		{
			throw new NotSupportedException();
		}
		public override int EndRead(IAsyncResult asyncResult)
		{
			return this.stream.EndRead(asyncResult);
		}
		public override void EndWrite(IAsyncResult asyncResult)
		{
			this.stream.EndWrite(asyncResult);
		}
		public override void Flush()
		{
			this.stream.Flush();
		}
		public override object InitializeLifetimeService()
		{
			throw new NotSupportedException();
		}
		public override long Length
		{
			get { return this.stream.Length; }
		}
		public override long Position
		{
			get { return this.stream.Position; }
			set { this.stream.Position = value; }
		}
		public override int Read(byte[] buffer, int offset, int count)
		{
			return this.stream.Read(buffer, offset, count);
		}
		public override int ReadByte()
		{
			return this.stream.ReadByte();
		}
		public override long Seek(long offset, SeekOrigin origin)
		{
			return this.stream.Seek(offset, origin);
		}
		public override void SetLength(long value)
		{
			this.stream.SetLength(value);
		}
		public override void Write(byte[] buffer, int offset, int count)
		{
			this.stream.Write(buffer, offset, count);
		}
		public override void WriteByte(byte value)
		{
			this.stream.WriteByte(value);
		}

		public IndisposableStream(Stream stream)
		{
			if (stream == null)
				throw new ArgumentNullException("stream");

			this.stream = stream;
		}

		private readonly Stream stream;
	}
}