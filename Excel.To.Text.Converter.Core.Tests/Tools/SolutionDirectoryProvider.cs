using System;
using System.IO;
using System.Reflection;

namespace Excel.To.Text.Converter.Core.Tests.Tools
{
	public static class SolutionDirectoryProvider
	{
		public static DirectoryInfo Get(string solutionFilename)
		{
			var assemblyDirectory = GetAssemblyDirectory();
			var currentDirectory = GetCurrentDirectory();
			var solutionDirectory = FindSolutionDirectory(assemblyDirectory, solutionFilename) ?? FindSolutionDirectory(currentDirectory, solutionFilename);

			if(solutionDirectory == null)
				throw new Exception($"Не получилось найти директорию с файлом '{solutionFilename}'. Текущая директория '{currentDirectory.FullName}'. Директория со сборкой '{assemblyDirectory.FullName}'");

			return solutionDirectory;
		}

		private static DirectoryInfo GetAssemblyDirectory()
		{
			return new FileInfo(Assembly.GetExecutingAssembly().Location).Directory;
		}

		private static DirectoryInfo GetCurrentDirectory()
		{
			return new DirectoryInfo(Directory.GetCurrentDirectory());
		}

		private static DirectoryInfo FindSolutionDirectory(DirectoryInfo directory, string solutionFilename)
		{
			DirectoryInfo parentDirectory;
			while(!directory.IsRoot())
			{
				if(directory.Contains(solutionFilename))
					return directory;

				parentDirectory = directory.Parent;
				if(parentDirectory == null)
					return null;
				directory = parentDirectory;
			}

			return null;
		}
	}
}