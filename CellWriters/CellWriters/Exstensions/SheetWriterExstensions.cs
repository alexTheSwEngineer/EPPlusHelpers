using CellWriters.Exstensions;
using OfficeOpenXml;
using System;
using System.Drawing;

namespace CellWriter.EPPlusHelpers.Excell
{
    public static class SheetWriterExstensions
    {
        public static void WithColor(this ISheetWriter writer, Color color, Action<ISheetWriter> writeAction)
        {
            writer.AsMonad()
                  .With(Modifiers.BgColor(color))
                  .Write(writeAction);
        }

        public static void WithSize(this ISheetWriter writer, float size, Action<ISheetWriter> writeAction)
        {
            writer.AsMonad()
                  .With(Modifiers.FontSize(size))
                  .Write(writeAction);
        }

        

    }
}