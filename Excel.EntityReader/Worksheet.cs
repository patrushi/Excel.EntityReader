using System;
using System.Collections.Generic;
using System.Linq;
using Excel.EntityReader.Columns;
using Excel.EntityReader.Exceptions;
using OfficeOpenXml;

namespace Excel.EntityReader
{
    public abstract class Worksheet
    {
        public bool ReadOnly { get; set; }

        public string Name { get; }

        public IList<ColumnBase> Columns { get; private set; }

        public ColumnBase ValueColumn { get; protected set; }

        public ColumnBase KeyColumn { get; protected set; }

        protected internal virtual int HeaderRowsCount { get; private set; }

        protected internal virtual Type EntityType => null;

        public ExcelReadingWorksheetException ExcelReadingWorksheetException;

        protected Worksheet(string name)
        {
            Name = name;
            Columns = new List<ColumnBase>();
            HeaderRowsCount = 1;
        }

        #region Запись в Excel

        protected internal virtual void WriteToExcel(ExcelWorksheet excelWorksheet)
        {
            WriteStructureToExcel(excelWorksheet);
            WriteDataToExcel(excelWorksheet);
            excelWorksheet.View.FreezePanes(HeaderRowsCount + 1, 1);
            excelWorksheet.Cells[excelWorksheet.Dimension.Address].AutoFitColumns();
        }

        protected abstract void WriteStructureToExcel(ExcelWorksheet excelWorksheet);

        protected abstract void WriteDataToExcel(ExcelWorksheet excelWorksheet);

        #endregion

        #region Чтение из Excel

        protected internal virtual void ReadFromExcel(ExcelWorksheet excelWorksheet, bool ignoreErrorRows = false)
        {
            if (ReadOnly) return;
            ExcelReadingWorksheetException = null;
            foreach (var columnBase in Columns)
            {
                columnBase.BeforeReadAll();
            }
            ReadDataFromExcel(excelWorksheet);

            if (!ignoreErrorRows && ExcelReadingWorksheetException != null)
            {
                throw new InvalidOperationException("Ошибка чтения Excel-файла.\n" + ExcelReadingWorksheetException.Message); 
            }
        }

        protected abstract void ReadDataFromExcel(ExcelWorksheet excelWorksheet);

        // определяет значение ключевой колонки по значению колонки значения
        protected internal abstract object GetKeyByValue(object value);

        protected internal abstract object GetValueByKey(object value);

        #endregion

        public abstract Worksheet Clone(string newName = null);
    }
}