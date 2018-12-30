namespace Excel.To.Text.Converter.Core.Formatters
{
	public interface ITextFormatter
	{
		string Format(string text, int width);
	}
}