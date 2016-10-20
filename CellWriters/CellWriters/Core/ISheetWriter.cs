using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace CellWriter.EPPlusHelpers.Excell
{
    public interface ISheetWriter
    {
        /// <summary>
        /// Writes all objects to a separate cell. Moves the cursor to the next one. Subsequent calls will write in the same row.
        /// </summary>
        /// <param name="values"></param>
        void Write(params object[] values);

        /// <summary>
        /// Writes all objects to a separate cell in the same row as the cursor and then moves it to the next row.
        /// </summary>
        /// <param name="values"></param>
        void WriteLine(params object[] values);

        /// <summary>
        /// Applies all cell modifiers for all write/writeLine calls and executes the action.
        /// </summary>
        /// <param name="cellModifier"></param>
        /// <param name="action"></param>
        void With(IEnumerable<Action<ExcelRange>> cellModifier, Action<ISheetWriter> action);

        /// <summary>
        /// Creates a <see cref="SheetWriterMonad"/> for this ISheetWriter
        /// </summary>
        /// <returns></returns>
        SheetWriterMonad AsMonad();
    }
}