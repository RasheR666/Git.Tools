using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using ApprovalTests;
using ApprovalTests.Namers;
using ApprovalTests.Reporters;
using Excel.To.Text.Converter.Core.Formatters;
using Excel.To.Text.Converter.Core.Tests.Tools;
using Excel.To.Text.Converter.Core.Xlsx;
using NUnit.Framework;

namespace Excel.To.Text.Converter.Core.Tests
{
	[TestFixture]
	[UseReporter(typeof(DiffReporter))]
	[UseApprovalSubdirectory("XlsxToTextConversionTest")]
	public class XlsxToTextConversionTest
	{
		[Test]
		[TestCaseSource(nameof(XlsxFileTestCases))]
		public void Convert(FileInfo xlsxFile)
		{
			using(ApprovalResults.ForScenario(xlsxFile.Name))
			{
				var sb = new StringBuilder();
				var xlsxDocument = new XlsxDocumentReader().Read(xlsxFile);
				new XlsxDocumentPrinter(new CenterAlignmentTextFormatter(), new StringBuilderWriter(sb)).Print(xlsxDocument);
				Approvals.Verify(sb.ToString());
			}
		}

		private static IEnumerable<TestCaseData> XlsxFileTestCases => XlsxFilesDirectory.GetFiles("*").Select(xlsxFile => new TestCaseData(xlsxFile).SetName(xlsxFile.Name));

		private static DirectoryInfo XlsxFilesDirectory => TestingProjectDirectoryProvider.ExcelToTextConverterCoreTestsDirectory.GetOrCreateSubDirectory("XlsxFiles");
	}
}