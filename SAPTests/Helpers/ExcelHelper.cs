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

    public static void ResetTestResults(string dataFilePath, string sheet = "")
    {
        using (XLWorkbook workbook = new XLWorkbook(dataFilePath))
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

    public static void UpdateTestResult(string dataFilePath, string sheet, string name, string result)
    {
        using (var workbook = new XLWorkbook(dataFilePath))
        {
            IXLWorksheet worksheet = workbook.Worksheet(sheet);
            var resultColumn = worksheet.Row(1).CellsUsed().FirstOrDefault(c => c.GetValue<string>().Equals("result", StringComparison.OrdinalIgnoreCase))?.Address.ColumnNumber;

            if (resultColumn == null)
            {
                throw new InvalidOperationException("Result column not found.");
            }

            // Skip header row
            IEnumerable<IXLRow> rows = worksheet.RowsUsed().Skip(1);
            foreach (IXLRow row in rows)
            {
                // Assuming the name is in column C
                if (string.Equals(row.Cell(3).GetValue<string>().Trim(), name, StringComparison.OrdinalIgnoreCase))
                {
                    row.Cell(resultColumn.Value).Value = result;
                    break;
                }
            }
            workbook.Save();
        }
    }
}
