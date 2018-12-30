using System;
using System.Collections.Generic;
using System.IO;
using ExcelDataReader;

namespace Excel.To.Text.Converter.Core.Xlsx
{
	public class XlsxDocumentReader
	{
		public XlsxDocument Read(FileInfo xlsxFile)
		{
			using(var fileStream = File.Open(xlsxFile.FullName, FileMode.Open, FileAccess.Read))
			{
				using(var reader = ExcelReaderFactory.CreateOpenXmlReader(fileStream))
				{
					return ReadXlsxDocument(xlsxFile, reader);
				}
			}
		}

		private XlsxDocument ReadXlsxDocument(FileInfo xlsxFile, IExcelDataReader reader)
		{
			var tables = new List<Table>();
			do
			{
				tables.Add(ReadTable(reader));
			} while(reader.NextResult()); // get next sheet

			return new XlsxDocument
			{
				Name = xlsxFile.Name,
				Tables = tables.ToArray()
			};
		}

		private Table ReadTable(IExcelDataReader reader)
		{
			var table = new Table
			{
				Name = reader.Name,
				Cells = new string[reader.RowCount, reader.FieldCount],
				ColumnWidth = new int[reader.FieldCount]
			};

			for(var rowIndex = 0; reader.Read(); rowIndex++)
			{
				for(var columnIndex = 0; columnIndex < reader.FieldCount; columnIndex++)
				{
					var cellValue = reader.GetValue(columnIndex)?.ToString() ?? string.Empty;
					table.Cells[rowIndex, columnIndex] = cellValue;
					table.ColumnWidth[columnIndex] = Math.Max(table.ColumnWidth[columnIndex], cellValue.Length);
				}
			}

			return table;
		}
	}
}