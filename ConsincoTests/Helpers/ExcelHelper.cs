using ClosedXML.Excel;

public class ExcelHelper
{
    private string _filePath;
    private string _sheetName;
    private IXLWorksheet _worksheet;

    public ExcelHelper(string filePath, string sheetName)
    {
        _filePath = filePath;
        _sheetName = sheetName;
    }

    public static void ResetTestResults(string dataSetFilePath, string sheet = "")
    {
        using (var workbook = new XLWorkbook(dataSetFilePath))
        {
            // Determine the sheets to process
            var sheets = string.IsNullOrEmpty(sheet)
                ? workbook.Worksheets.Where(ws => !ws.Name.Equals("Config", StringComparison.OrdinalIgnoreCase)).Select(ws => ws.Name).ToList()
                : new List<string> { sheet };

            foreach (var sheetName in sheets)
            {
                var worksheet = workbook.Worksheet(sheetName);
                var resultColumn = worksheet.Row(1).CellsUsed()
                    .FirstOrDefault(c => c.GetValue<string>().Equals("result", StringComparison.OrdinalIgnoreCase))?.Address.ColumnNumber;

                if (resultColumn == null)
                {
                    throw new InvalidOperationException($"Result column not found in sheet '{sheetName}'.");
                }

                // Skip header row and update all rows
                var rows = worksheet.RowsUsed().Skip(1);
                foreach (var row in rows)
                {
                    row.Cell(resultColumn.Value).Value = "";
                }
            }
            workbook.Save();
        }
    }

    public static void UpdateTestResult(string dataSetFilePath, string sheet, int testId, string result)
    {
        using (var workbook = new XLWorkbook(dataSetFilePath))
        {
            var worksheet = workbook.Worksheet(sheet);
            var resultColumn = worksheet.Row(1).CellsUsed().FirstOrDefault(c => c.GetValue<string>().Equals("result", StringComparison.OrdinalIgnoreCase))?.Address.ColumnNumber;

            if (resultColumn == null)
            {
                throw new InvalidOperationException("Result column not found.");
            }

            // Skip header row
            var rows = worksheet.RowsUsed().Skip(1);
            foreach (var row in rows)
            {
                // Assuming the testId is in column B
                if (int.TryParse(row.Cell(2).GetValue<string>().Trim(), out int cellTestId) && cellTestId == testId)
                {
                    row.Cell(resultColumn.Value).Value = result;
                    break;
                }
            }
            workbook.Save();
        }
    }
}
