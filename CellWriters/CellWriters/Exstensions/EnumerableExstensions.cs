using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
