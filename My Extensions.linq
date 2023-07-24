<Query Kind="Program">
  <NuGetReference>ClosedXML</NuGetReference>
  <Namespace>ClosedXML.Excel</Namespace>
</Query>

void Main()
{
	var sampleFile = @"C:\Users\paulp\Documents\GitHub\Linqpad-Extensions\SampleData.xlsx";
	
	var table = MyExtensions.ReadRange(sampleFile, "Sheet1");
	
	table.Dump();
	
	const string testSheet = "test";
	MyExtensions.AddSheet(sampleFile, testSheet);
	MyExtensions.WriteRange(sampleFile, testSheet, table);
	MyExtensions.ReadRange(sampleFile, testSheet).Dump();
	MyExtensions.DeleteSheet(sampleFile, testSheet);
}

public static class MyExtensions
{
	public static DataTable	ReadRange(string filePath, string sheetName)
	{
		using (var wb = new XLWorkbook(filePath))
		{
			var ws = wb.Worksheet(sheetName);
			var firstCell = ws.FirstCellUsed();
			var lastCell = ws.LastCellUsed();
			var range = ws.Range(firstCell.Address, lastCell.Address);
			return range.CreateTable().AsNativeDataTable();
		}
	}
	
	public static void WriteRange(string filePath, string sheetName, DataTable table)
	{
		using (var wb = new XLWorkbook(filePath))
		{
			var ws = wb.Worksheet(sheetName);
			ws.Cell(1,1).InsertTable(table);
			wb.Save();
		}
	}
	
	internal static void DeleteSheet(string filePath, string sheetName)
	{
		using (var wb = new XLWorkbook(filePath))
		{
			wb.Worksheets.Delete(sheetName);
			wb.Save();
		}
	}

	internal static void AddSheet(string filePath, string sheetName)
	{
		using (var wb = new XLWorkbook(filePath))
		{
			wb.Worksheets.Add(sheetName);
			wb.Save();
		}
	}
}




