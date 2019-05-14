using System;
using Excel.EntityReader.Exceptions;
using OfficeOpenXml;

namespace Excel.EntityReader.Columns
{
    public abstract class ColumnBase
    {
        public const int MAX_ROW_NUMBER = 1048576;

        public readonly string Name;

        public bool IsRequired { get; set; }

        public int ColumnNumber
        {
            get { return _columnNumber; }
            set
            {
                _columnNumber = value;
                ColumnCode = GetColumnCode(_columnNumber);
            }
        }

        private int _columnNumber;

        public string ColumnCode { get; private set; }

        public Worksheet Worksheet { protected set; get; }

        protected ColumnBase(string name)
        {
            Name = name;
        }

        protected internal void SetWorksheet(Worksheet worksheet, int columnIndex)
        {
            Worksheet = worksheet;
            if (ColumnNumber == 0) ColumnNumber = columnIndex;
        }

        public static string GetColumnCode(int columnNumber)
        {
            var dividend = columnNumber;
            var columnName = string.Empty;

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        public string GetColumnDataRangeForValidation()
        {
            var columnCode = GetColumnCode(ColumnNumber);
            return $"${columnCode}${Worksheet.HeaderRowsCount + 1}:${columnCode}${MAX_ROW_NUMBER}";
        }

        protected internal virtual void WriteStructureToExcel(ExcelWorksheet excelWorksheet)
        {
            // заголовок
            excelWorksheet.Cells[1, ColumnNumber].Value = Name;

            // валидация данных
            WriteDataValidation(excelWorksheet);
        }

        protected virtual void WriteDataValidation(ExcelWorksheet excelWorksheet)
        { }

        protected internal virtual void WriteValue(ExcelWorksheet excelWorksheet, int rowNumber, object value)
        {
            WriteRawValue(excelWorksheet, rowNumber, value);
        }

        protected internal virtual void WriteRawValue(ExcelWorksheet excelWorksheet, int rowNumber, object value)
        {
            excelWorksheet.Cells[rowNumber, ColumnNumber].Value = value;
        }

        public virtual object ReadValue(ExcelWorksheet excelWorksheet, int rowNumber)
        {
            var value = excelWorksheet.Cells[rowNumber, ColumnNumber].Value;
            if (IsRequired && value == null)
            {
                throw new ExcelReadingCellException(this, rowNumber, "Пустое значение не допустимо.");
            }
            return value;
        }

        public abstract ColumnBase Clone(string newName = null);

        public virtual void BeforeReadAll()
        {
        }
    }
}