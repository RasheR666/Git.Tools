using System;
using System.IO;
using System.Text;

namespace Excel.To.Text.Converter.Core.Writers
{
	public class FileWriter : IWriter, IDisposable
	{
		public FileWriter(string filepath)
		{
			writer = new StreamWriter(filepath, false, Encoding.UTF8);
		}

		public void Dispose()
		{
			writer.Dispose();
		}

		public void Write(string text = null)
		{
			writer.Write(text);
		}

		public void WriteLine(string text = null)
		{
			writer.WriteLine(text);
		}

		private readonly StreamWriter writer;
	}
}