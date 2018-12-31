using System.IO;
using System.Linq;

namespace Excel.To.Text.Converter.Core.Tests.Tools
{
	public static class DirectoryInfoExtensions
	{
		public static bool IsRoot(this DirectoryInfo directory)
		{
			return Path.GetPathRoot(directory.FullName) == directory.FullName;
		}

		public static bool Contains(this DirectoryInfo directory, string filename)
		{
			return directory.GetFiles().Any(x => x.Name == filename);
		}

		public static DirectoryInfo GetOrCreateSubDirectory(this DirectoryInfo directory, string directoryName)
		{
			var subDirectory = directory.GetDirectories().FirstOrDefault(x => x.Name == directoryName);
			return subDirectory ?? directory.CreateSubdirectory(directoryName);
		}
	}
}