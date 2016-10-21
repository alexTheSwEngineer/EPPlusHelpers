namespace CellWriters.Exstensions
{
    public static class ObjectArrayExstensions
    {
        public static bool Empty(this object[] values)
        {
            return values == null || values.Length <= 0;
        }

        public static bool Any(this object[] values)
        {
            return !values.Empty();
        }
    }
}
