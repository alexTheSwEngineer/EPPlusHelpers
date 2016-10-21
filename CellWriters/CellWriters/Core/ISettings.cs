using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace CellWriters.Core
{
    public interface ISettings
    {
         IEnumerable<Action<ExcelRange>> Modifiers { get; }
         ISettings ApplyTo(ExcelRange cell);
         ISettings With(Action<ExcelRange> modifier);
         ISettings With(ISettings other);
    }         
}
