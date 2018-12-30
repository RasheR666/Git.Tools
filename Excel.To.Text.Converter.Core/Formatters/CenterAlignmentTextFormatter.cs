using System.Text;

namespace Excel.To.Text.Converter.Core.Formatters
{
	public class CenterAlignmentTextFormatter : ITextFormatter
	{
		public string Format(string text, int width)
		{
			if(string.IsNullOrWhiteSpace(text))
				return new string(' ', width);

			var whitespacesCount = width - text.Length;
			var leftIndentationSize = whitespacesCount / 2;
			var rightIndentationSize = whitespacesCount - leftIndentationSize;

			var sb = new StringBuilder();
			sb.Append(' ', leftIndentationSize);
			sb.Append(text);
			sb.Append(' ', rightIndentationSize);
			return sb.ToString();
		}
	}
}