using Excel.EntityReader.Columns;

namespace Excel.EntityReader.Exceptions
{
    public class ExcelReadingCellException : ExcelException
    {
        private readonly ColumnBase _columnBase;
        private readonly long _row;
        private readonly string _message;

        public ExcelReadingCellException(ColumnBase columnBase, long row, string message)
        {
            _columnBase = columnBase;
            _row = row;
            _message = message;
        }

        public override string Message => $"Ошибка чтения. Лист: [{_columnBase.Worksheet.Name}], колонка: [{_columnBase.Name} ({_columnBase.ColumnCode})], строка {_row}. {_message}";
    }
}