using System.IO;

namespace Excel.To.Text.Converter.Core.Tests.Tools
{
	public static class TestingProjectDirectoryProvider
	{
		public static DirectoryInfo SolutionDirectory { get; } = SolutionDirectoryProvider.Get("Git.Tools.sln");
		public static DirectoryInfo ExcelToTextConverterCoreTestsDirectory { get; } = SolutionDirectory.GetOrCreateSubDirectory("Excel.To.Text.Converter.Core.Tests");
	}
}