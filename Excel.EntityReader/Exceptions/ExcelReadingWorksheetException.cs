using System.Collections.Generic;
using System.Linq;

namespace Excel.EntityReader.Exceptions
{
    public class ExcelReadingWorksheetException : ExcelException
    {
        public IList<ExcelReadingCellException> ExcelReadingCellExceptions = new List<ExcelReadingCellException>();

        public override string Message => string.Join("\n", ExcelReadingCellExceptions.Select(e => e.Message));
    }
}