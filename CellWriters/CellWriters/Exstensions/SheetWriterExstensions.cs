using CellWriters.Exstensions;
using OfficeOpenXml;
using System;
using System.Drawing;

namespace CellWriter.EPPlusHelpers.Excell
{
    public static class SheetWriterExstensions
    {
        public static ISheetWriter WithColor(this ISheetWriter writer, Color color, Action<ISheetWriter> writeAction)
        {
            var settings = SettingsExstensions.BgColor(color);
            return writer.WithAddedSettings(settings,writeAction);
        }

        public static ISheetWriter WithSize(this ISheetWriter writer, float size, Action<ISheetWriter> writeAction)
        {
            var settings = SettingsExstensions.FontSize(size);
            return writer.WithAddedSettings(settings, writeAction);
        }
    }
}