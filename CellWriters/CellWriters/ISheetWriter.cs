using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace CellWriter.EPPlusHelpers.Excell
{
    public interface ISheetWriter
    {
        void Write(params object[] values);
        void WriteLine(params object[] values);
        void With(IEnumerable<Action<ExcelRange>> cellModifier, Action<ISheetWriter> action);
        SheetWriterMonad AsMonad();
    }
}