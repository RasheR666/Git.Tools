using System;
using System.IO;
using System.Text;
using Excel.To.Text.Converter.Core.Formatters;
using Excel.To.Text.Converter.Core.Writers;
using Excel.To.Text.Converter.Core.Xlsx;

namespace Excel.To.Text.Converter
{
	internal class Program
	{
		private static void Main(string[] args)
		{
			if(!IsValid(args))
			{
				Console.WriteLine($"args is not valid. args: [{string.Join(", ", args)}]");
				return;
			}
			Console.OutputEncoding = Encoding.UTF8;
			Convert(args[0]);
		}

		private static bool IsValid(string[] args)
		{
			return args != null && args.Length == 1 && File.Exists(args[0]);
		}

		private static void Convert(string filepath)
		{
			try
			{
				var inputFile = new FileInfo(filepath);
				var xlsxDocument = new XlsxDocumentReader().Read(inputFile);
				new XlsxDocumentPrinter(new CenterAlignmentTextFormatter(), new ConsoleWriter()).Print(xlsxDocument);
			}
			catch(Exception exception)
			{
				Console.WriteLine($"Exception: {exception}");
			}
		}
	}
}