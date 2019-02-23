using System;
using System.Diagnostics;
using System.IO;
using Excel.To.Text.Converter.Core.Formatters;
using Excel.To.Text.Converter.Core.Writers;
using Excel.To.Text.Converter.Core.Xlsx;

namespace Diff.Tool
{
	internal class Program
	{
		private static void Main(string[] args)
		{
			var beforeFilepath = args[0];
			var afterFilepath = args[1];

			var beforeFileExtension = GetExtension(beforeFilepath);
			var afterFileExtension = GetExtension(afterFilepath);

			if(beforeFileExtension == xlsx && afterFileExtension == xlsx)
			{
				ShowXlsxDiff(beforeFilepath, afterFilepath);
			}
			else if(beforeFileExtension == docx && afterFileExtension == docx)
			{
				ShowDocxDiff(beforeFilepath, afterFilepath);
			}
			else
			{
				ShowDiff(beforeFilepath, afterFilepath);
			}
		}

		private static string GetExtension(string filepath)
		{
			return Path.GetExtension(filepath);
		}

		private static void ShowXlsxDiff(string beforeFilepath, string afterFilepath)
		{
			var dateTime = DateTime.Now;
			var beforeTextFile = GetFilePathInTempDirectory($"before.{ToString(dateTime)}.xlsx.txt");
			var afterTextFile = GetFilePathInTempDirectory($"after.{ToString(dateTime)}.xlsx.txt");

			ConvertXlsxToTextFile(beforeFilepath, beforeTextFile);
			ConvertXlsxToTextFile(afterFilepath, afterTextFile);

			TortoiseGitMerge.Compare(beforeTextFile, afterTextFile);
		}

		private static void ConvertXlsxToTextFile(string xlsxFilepath, string txtFilePath)
		{
			using(var outputStream = new FileWriter(txtFilePath))
			{
				var inputFile = new FileInfo(xlsxFilepath);
				var xlsxDocument = new XlsxDocumentReader().Read(inputFile);
				new XlsxDocumentPrinter(new CenterAlignmentTextFormatter(), outputStream).Print(xlsxDocument);
			}
		}

		private static void ShowDocxDiff(string beforeFilepath, string afterFilepath)
		{
			var dateTime = DateTime.Now;
			var beforeTextFile = GetFilePathInTempDirectory($"before.{ToString(dateTime)}.docx.txt");
			var afterTextFile = GetFilePathInTempDirectory($"after.{ToString(dateTime)}.docx.txt");

			ConvertDocxToTextFile(beforeFilepath, beforeTextFile);
			ConvertDocxToTextFile(afterFilepath, afterTextFile);

			TortoiseGitMerge.Compare(beforeTextFile, afterTextFile);
		}

		private static void ConvertDocxToTextFile(string docxFilepath, string txtFilePath)
		{
			var pandocExePath = Pandoc.GetExePath();
			var process = Process.Start(pandocExePath, $"-s \"{docxFilepath}\" -t markdown -o \"{txtFilePath}\"");
			process?.WaitForExit();
		}

		private static void ShowDiff(string beforeFilepath, string afterFilepath)
		{
			var dateTime = DateTime.Now;
			var beforeFile = GetFilePathInTempDirectory($"before.{ToString(dateTime)}{Path.GetExtension(beforeFilepath)}");
			var afterFile = GetFilePathInTempDirectory($"after.{ToString(dateTime)}{Path.GetExtension(afterFilepath)}");

			File.Copy(beforeFilepath, beforeFile);
			File.Copy(afterFilepath, afterFile);

			TortoiseGitMerge.Compare(beforeFile, afterFile);
		}

		private static string GetFilePathInTempDirectory(string filename)
		{
			return Path.Combine(Path.GetTempPath(), filename);
		}

		private static string ToString(DateTime dateTime)
		{
			return dateTime.ToString("yyyy-MM-dd HH-mm-ss");
		}

		private const string xlsx = ".xlsx";
		private const string docx = ".docx";
	}
}