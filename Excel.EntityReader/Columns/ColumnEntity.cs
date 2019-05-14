using Excel.EntityReader.Exceptions;
using OfficeOpenXml;

namespace Excel.EntityReader.Columns
{
    public class ColumnEntity : ColumnBase
    {
        protected readonly Worksheet ReferencedWorksheet;

        public ColumnEntity(Worksheet referencedWorksheet, string name = null) : base(name ?? referencedWorksheet.Name)
        {
            ReferencedWorksheet = referencedWorksheet;
        }

        protected override void WriteDataValidation(ExcelWorksheet excelWorksheet)
        {
            base.WriteDataValidation(excelWorksheet);

            //Add a List validation to the columnNumber
            var val = excelWorksheet.DataValidations.AddListValidation(GetColumnDataRangeForValidation());
            //Define the Cells with the accepted values
            val.Formula.ExcelFormula =
                $"='{ReferencedWorksheet.Name}'!{ReferencedWorksheet.ValueColumn.GetColumnDataRangeForValidation()}";
            val.ShowErrorMessage = true;
            //Style of warning. "information" and "warning" allow users to ignore the validation,
            //while "stop" and "undefined" doesn't
            val.ErrorStyle = OfficeOpenXml.DataValidation.ExcelDataValidationWarningStyle.stop;
            //Title of the error mesage box
            //val.ErrorTitle = "This is the title";
            //Message of the error
            //val.Error = "This is the message";
            //Set to true to show a prompt when user clics on the cell
            //val.ShowInputMessage = true;
            //Set the message for the prompt
            //val.Prompt = "This is a input message";
            //Set the title for the prompt
            //val.PromptTitle = "This is the title from the input message";
        }

        protected internal override void WriteValue(ExcelWorksheet excelWorksheet, int rowNumber, object value)
        {
            var rawValue = ReferencedWorksheet.GetValueByKey(value);
            WriteRawValue(excelWorksheet, rowNumber, rawValue);
        }

        public override object ReadValue(ExcelWorksheet excelWorksheet, int rowNumber)
        {
            var rawValue = base.ReadValue(excelWorksheet, rowNumber);

            if (rawValue == null) return null;

            var value = ReferencedWorksheet.GetKeyByValue(rawValue);
            if (IsRequired && value == null)
            {
                throw new ExcelReadingCellException(this, rowNumber, $"Значение [{rawValue}] не найдено на листе [{ReferencedWorksheet.Name}] в колонке [{ReferencedWorksheet.KeyColumn.Name}], в результате получено пустое значение, которое не допустимо.");
            }

            return value;
        }

        public override ColumnBase Clone(string newName = null)
        {
            return new ColumnEntity(ReferencedWorksheet, newName ?? Name);
        }
    }
}