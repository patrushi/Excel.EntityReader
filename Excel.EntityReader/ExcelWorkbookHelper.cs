using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;

namespace Excel.EntityReader
{
    public static class ExcelWorkbookHelper
    {
        public static IList<string> GetRawRow(ExcelWorkbook excelWorkbook, string worksheetName, int rowNumber)
        {
            var excelWorksheet = excelWorkbook.Worksheets[worksheetName];
            var columnsCount = excelWorksheet.Dimension.Columns;
            return Enumerable.Range(1, columnsCount)
                .Select(columnNumber => excelWorksheet.Cells[rowNumber, columnNumber].Value?.ToString()).ToList();
        }
    }
}