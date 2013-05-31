namespace ExcelStream
{
	using System;
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
	}
}