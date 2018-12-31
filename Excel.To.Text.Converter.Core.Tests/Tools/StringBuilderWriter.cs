using System.Text;
using Excel.To.Text.Converter.Core.Writers;

namespace Excel.To.Text.Converter.Core.Tests.Tools
{
	public class StringBuilderWriter : IWriter
	{
		public StringBuilderWriter(StringBuilder stringBuilder)
		{
			this.stringBuilder = stringBuilder;
		}

		public void Write(string text = null)
		{
			stringBuilder.Append(text);
		}

		public void WriteLine(string text = null)
		{
			stringBuilder.AppendLine(text);
		}

		private readonly StringBuilder stringBuilder;
	}
}