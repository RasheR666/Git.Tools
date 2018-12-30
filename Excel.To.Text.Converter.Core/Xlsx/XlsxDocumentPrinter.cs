using System;
using Excel.To.Text.Converter.Core.Formatters;
using Excel.To.Text.Converter.Core.Writers;

namespace Excel.To.Text.Converter.Core.Xlsx
{
	public class XlsxDocumentPrinter
	{
		public XlsxDocumentPrinter(ITextFormatter textFormatter, IWriter writer)
		{
			this.textFormatter = textFormatter;
			this.writer = writer;
		}

		public void Print(XlsxDocument xlsxDocument)
		{
			foreach(var table in xlsxDocument.Tables)
			{
				PrintBorder(table);
				for(var rowIndex = 0; rowIndex < table.Cells.GetLength(0); rowIndex++)
				{
					var tmpRowIndex = rowIndex;
					PrintRow(table, "|", (columnIndex, columnWidth) => textFormatter.Format(table.Cells[tmpRowIndex, columnIndex], columnWidth));
					PrintBorder(table);
				}
				writer.WriteLine();
			}
		}

		private void PrintBorder(Table table) => PrintRow(table, "+", (columnIndex, columnWidth) => new string('-', columnWidth));

		private void PrintRow(Table table, string columnBorderChar, Func<int, int, string> getColumnText)
		{
			writer.Write(columnBorderChar);
			for(var columnIndex = 0; columnIndex < table.Cells.GetLength(1); columnIndex++)
			{
				writer.Write(getColumnText(columnIndex, GetColumnWidth(table, columnIndex)));
				writer.Write(columnBorderChar);
			}
			writer.WriteLine();
		}

		private int GetColumnWidth(Table table, int columnIndex) => table.ColumnWidth[columnIndex] + 2;

		private readonly ITextFormatter textFormatter;
		private readonly IWriter writer;
	}
}