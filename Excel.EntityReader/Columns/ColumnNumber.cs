using OfficeOpenXml;
using OfficeOpenXml.DataValidation;

namespace Excel.EntityReader.Columns
{
    public class ColumnNumber : ColumnBase
    {
        public ColumnNumber(string name) : base(name)
        {
        }

        protected override void WriteDataValidation(ExcelWorksheet excelWorksheet)
        {
            base.WriteDataValidation(excelWorksheet);

            var val = excelWorksheet.DataValidations.AddDecimalValidation(GetColumnDataRangeForValidation());
            val.Operator = ExcelDataValidationOperator.greaterThanOrEqual;
            val.Formula.Value = 0;
            val.ShowErrorMessage = true;
            val.ErrorStyle = ExcelDataValidationWarningStyle.stop;
        }

        public override ColumnBase Clone(string newName = null)
        {
            return new ColumnNumber(newName ?? Name);
        }
    }
}