namespace Excel.To.Text.Converter.Core.Xlsx
{
	public class XlsxDocument
	{
		public string Name { get; set; }

		public Table[] Tables { get; set; }
	}

	public class Table
	{
		public string Name { get; set; }

		public string[,] Cells { get; set; }

		public int[] ColumnWidth { get; set; }
	}
}