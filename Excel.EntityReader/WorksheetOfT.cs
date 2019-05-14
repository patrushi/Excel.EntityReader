using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using Excel.EntityReader.Columns;
using Excel.EntityReader.Exceptions;
using OfficeOpenXml;

namespace Excel.EntityReader
{
    public class Worksheet<T> : Worksheet
        where T : new()
    {
        protected internal override Type EntityType => typeof(T);

        protected readonly IDictionary<ColumnBase, Func<T, object>> GetterFuncs =
            new ConcurrentDictionary<ColumnBase, Func<T, object>>();

        protected readonly IDictionary<ColumnBase, Action<T, object>> SetterFunc =
            new ConcurrentDictionary<ColumnBase, Action<T, object>>();

        protected readonly Func<T> EntityFactory;

        public IEnumerable<T> EntityList => _entityList;

        private readonly IList<T> _entityList = new List<T>();

        /// <summary></summary>
        /// <param name="name">Название</param>
        /// <param name="entityFactory">Фабрика по созданию новой сущности, по умолчанию просто new T()</param>
        public Worksheet(string name, Func<T> entityFactory = null) : base(name)
        {
            EntityFactory = entityFactory ?? (() => new T());
        }

        public Worksheet<T> AddColumn(ColumnBase column, Func<T, object> getterFunc, Action<T, object> setterFunc)
        {
            Columns.Add(column);

            column.SetWorksheet(this, Columns.Count);

            GetterFuncs.Add(column, getterFunc);
            SetterFunc.Add(column, setterFunc);

            return this;
        }

        public Worksheet<T> AddValueColumn(ColumnBase column, Func<T, object> getterFunc, Action<T, object> setterFunc)
        {
            ValueColumn = column;

            return AddColumn(column, getterFunc, setterFunc);
        }

        public Worksheet<T> AddKeyColumn(ColumnBase column, Func<T, object> getterFunc, Action<T, object> setterFunc)
        {
            KeyColumn = column;

            return AddColumn(column, getterFunc, setterFunc);
        }

        public virtual void AddEntity(T entity)
        {
            _entityList.Add(entity);
        }

        public virtual void AddEntities(IEnumerable<T> entities)
        {
            foreach (var entity in entities)
            {
                AddEntity(entity);
            }
        }

        protected override void WriteStructureToExcel(ExcelWorksheet excelWorksheet)
        {
            foreach (var column in Columns)
            {
                column.WriteStructureToExcel(excelWorksheet);
            }
        }

        protected override void WriteDataToExcel(ExcelWorksheet excelWorksheet)
        {
            var rowNumber = HeaderRowsCount;
            foreach (var entity in _entityList)
            {
                rowNumber++;
                WriteEntityToExcel(excelWorksheet, rowNumber, entity);
            }
        }

        protected virtual void WriteEntityToExcel(ExcelWorksheet excelWorksheet, int rowNumber, T entity)
        {
            foreach (var getterFunc in GetterFuncs)
            {
                getterFunc.Key.WriteValue(excelWorksheet, rowNumber, getterFunc.Value(entity));
            }
        }

        protected override void ReadDataFromExcel(ExcelWorksheet excelWorksheet)
        {
            ExcelReadingWorksheetException = null;

            var rowsCount = excelWorksheet.Dimension.End.Row;

            for (var rowNumber = HeaderRowsCount + 1; rowNumber <= rowsCount; rowNumber++)
            {
                try
                {
                    var entity = ReadEntityFromExcel(excelWorksheet, rowNumber);
                    AddEntity(entity);
                }
                catch (ExcelException e)
                {
                    if (ExcelReadingWorksheetException == null)
                    {
                        ExcelReadingWorksheetException = new ExcelReadingWorksheetException();
                    }
                    ExcelReadingWorksheetException.ExcelReadingCellExceptions.Add((ExcelReadingCellException)e);
                }
            }
        }

        protected internal override object GetKeyByValue(object value)
        {
            return value == null
                ? null
                : GetterFuncs[KeyColumn](EntityList
                    .SingleOrDefault(e => Equals(GetterFuncs[ValueColumn](e).ToString(), value.ToString())));
        }

        protected internal override object GetValueByKey(object value)
        {
            return value == null
                ? null
                : GetterFuncs[ValueColumn](EntityList
                    .SingleOrDefault(e => Equals(GetterFuncs[KeyColumn](e).ToString(), value.ToString())));
        }

        #region Клонирование

        public override Worksheet Clone(string newName = null)
        {
            var worksheet = CloneCreateWorksheet(newName);
            worksheet.ReadOnly = ReadOnly;
            CloneCopyColumns(worksheet);
            return worksheet;
        }

        protected virtual Worksheet CloneCreateWorksheet(string newName)
        {
            return new Worksheet<T>(newName, EntityFactory);
        }

        protected virtual void CloneCopyColumns(Worksheet worksheet)
        {
            var ws = (Worksheet<T>) worksheet;
            foreach (var columnBase in Columns)
            {
                var newColumn = columnBase.Clone();
                if (columnBase == ValueColumn) ws.ValueColumn = newColumn;
                if (columnBase == KeyColumn) ws.KeyColumn = newColumn;
                ws.AddColumn(newColumn,
                    GetterFuncs[columnBase],
                    SetterFunc[columnBase]);
            }
        }

        #endregion

        protected virtual T ReadEntityFromExcel(ExcelWorksheet excelWorksheet, int rowNumber)
        {
            var entity = EntityFactory();
            foreach (var setterFunc in SetterFunc)
            {
                var value = setterFunc.Key.ReadValue(excelWorksheet, rowNumber);
                setterFunc.Value(entity, value);
            }
            return entity;
        }
    }
}