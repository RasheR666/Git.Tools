using System;

namespace Excel.To.Text.Converter.Core.Writers
{
	public class ConsoleWriter : IWriter
	{
		public void Write(string text = null)
		{
			Console.Write(text);
		}

		public void WriteLine(string text = null)
		{
			Console.WriteLine(text);
		}
	}
}