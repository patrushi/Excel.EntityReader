using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;

namespace Excel.EntityReader
{
    public class Workbook
    {
        private readonly IList<Worksheet> _worksheets;

        public Workbook()
        {
            _worksheets = new List<Worksheet>();
        }

        public Worksheet<T> AddWorksheet<T>(Worksheet<T> worksheet)
            where T : new()
        {
            _worksheets.Add(worksheet);
            return worksheet;
        }

        public Worksheet<T> GetEntityWorksheet<T>(string name = null)
            where T : new()
        {
            return name != null
                ? (Worksheet<T>) _worksheets.Single(w => w.Name == name)
                : (Worksheet<T>) _worksheets.Single(w => w.EntityType == typeof(T));
        }

        public PropertyWorksheet GetPropertyWorksheet(string name = null)
        {
            return (PropertyWorksheet) _worksheets.Single(w =>
                w.Name == (name ?? PropertyWorksheet.DEFAULT_PROPERTY_WORKSHEET_NAME));
        }

        #region Работа с Excel

        public void WriteToExcel(ExcelWorkbook excelWorkbook)
        {
            foreach (var worksheet in _worksheets)
            {
                var excelWorksheet = excelWorkbook.Worksheets.Add(worksheet.Name);
                worksheet.WriteToExcel(excelWorksheet);
            }
        }

        public void ReadFromExcel(ExcelWorkbook excelWorkbook, bool ignoreErrorRows = false)
        {
            // сначала обрабатываем справочники, чтобы загрузузились все ключи
            foreach (var worksheet in _worksheets.OrderBy(w => w is PropertyWorksheet ? 0 : 1))
            {
                var excelWorksheet = excelWorkbook.Worksheets.SingleOrDefault(w => w.Name == worksheet.Name);
                if (excelWorksheet == null) throw new ArgumentException(
                    $"Не найден лист с названием [{worksheet.Name}]");
                worksheet.ReadFromExcel(excelWorksheet, ignoreErrorRows);
            }
        }

        #endregion   
    }
}