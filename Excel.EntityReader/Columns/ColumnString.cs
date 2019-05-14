namespace Excel.EntityReader.Columns
{
    public class ColumnString : ColumnBase
    {
        public ColumnString(string name) : base(name)
        {
        }

        public override ColumnBase Clone(string newName = null)
        {
            return new ColumnString(newName ?? Name);
        }
    }
}