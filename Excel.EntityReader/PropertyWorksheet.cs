using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;

namespace Excel.EntityReader
{
    public class PropertyWorksheet : KeyValueWorksheet<KeyValuePair<object, object>>
    {
        public const string DEFAULT_PROPERTY_WORKSHEET_NAME = "Настройки";

        public PropertyWorksheet(string name = null) 
            : base(name ?? DEFAULT_PROPERTY_WORKSHEET_NAME, i => i.Key, i => i.Value)
        {
            ReadOnly = false;
        }

        // Конструктор для клонирования
        protected PropertyWorksheet(string name, bool ctorForClone)
            : base(name)
        {
        }

        protected override Worksheet CloneCreateWorksheet(string newName)
        {
            return new PropertyWorksheet(newName, true);
        }

        protected override KeyValuePair<object, object> ReadEntityFromExcel(ExcelWorksheet excelWorksheet, int rowNumber)
        {
            return new KeyValuePair<object, object>(KeyColumn.ReadValue(excelWorksheet, rowNumber).ToString(),
                ValueColumn.ReadValue(excelWorksheet, rowNumber));
        }

        public TValue GetProperty<TValue>(object key, Func<object, TValue> convertFunc)
        {
            return convertFunc(EntityList.Single(e => e.Key?.ToString() == key?.ToString()).Value);
        }

        public void SetProperty(object key, object value)
        {
            AddEntity(new KeyValuePair<object, object>(key, value));
        }

        public void AddProperties(IEnumerable<KeyValuePair<object, object>> properties)
        {
            foreach (var property in properties)
            {
                SetProperty(property.Key, property.Value);
            }
        }

        public void AddProperties(IDictionary<object, object> properties)
        {
            foreach (var property in properties)
            {
                SetProperty(property.Key, property.Value);
            }
        }
    }
}