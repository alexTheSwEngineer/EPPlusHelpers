using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CellWriters
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
