using System;
using Excel.EntityReader.Columns;

namespace Excel.EntityReader
{
    public class KeyValueWorksheet<T> : Worksheet<T>
        where T : new()
    {
        public KeyValueWorksheet(string name, Func<T, object> getterKeyFunc, Func<T, object> getterValueFunc) 
            : base(name)
        {
            AddKeyColumn(new ColumnString("Ключ"), getterKeyFunc, (p, v) => { });
            AddValueColumn(new ColumnString("Значение"), getterValueFunc, (p, v) => { });
        }

        // Конструктор для клонирования
        protected KeyValueWorksheet(string name)
            : base(name)
        {
        }

        protected override Worksheet CloneCreateWorksheet(string newName)
        {
            return new KeyValueWorksheet<T>(newName);
        }
    }
}