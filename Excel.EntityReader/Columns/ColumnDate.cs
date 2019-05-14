using System;
using Excel.EntityReader.Exceptions;
using OfficeOpenXml;

namespace Excel.EntityReader.Columns
{
    public class ColumnDate : ColumnBase
    {
        public ColumnDate(string name) : base(name)
        {
        }

        protected override void WriteDataValidation(ExcelWorksheet excelWorksheet)
        {
            base.WriteDataValidation(excelWorksheet);
            excelWorksheet.Column(ColumnNumber).Style.Numberformat.Format = "dd.mm.yyyy";
        }

        public override ColumnBase Clone(string newName = null)
        {
            return new ColumnDate(newName ?? Name);
        }

        public override object ReadValue(ExcelWorksheet excelWorksheet, int rowNumber)
        {
            var rawValue = excelWorksheet.Cells[rowNumber, ColumnNumber].Value?.ToString();

            var value = string.IsNullOrEmpty(rawValue)
                ? (DateTime?) null
                : DateTime.FromOADate(long.Parse(rawValue));

            if (IsRequired && value == null)
            {
                throw new ExcelReadingCellException(this, rowNumber, "Пустое значение не допустимо.");
            }

            return value;
        }
    }
}