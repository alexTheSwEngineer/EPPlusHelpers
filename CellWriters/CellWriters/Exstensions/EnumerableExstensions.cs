using System.Collections.Generic;

namespace CellWriters.Exstensions
{
    public static class EnumerableExstensions
    {
        public static IEnumerable<T> ToEnumerable<T>(this T obj)
        {
            return new[] { obj };
        }
    }
}
