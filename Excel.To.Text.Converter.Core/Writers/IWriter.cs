namespace Excel.To.Text.Converter.Core.Writers
{
	public interface IWriter
	{
		void Write(string text = null);
		void WriteLine(string text = null);
	}
}