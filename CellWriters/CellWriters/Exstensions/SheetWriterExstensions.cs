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
            var settings = SettingsExstensions.BgColor(color);
            writer.WithTempSettings(settings,writeAction);
        }

        public static void WithSize(this ISheetWriter writer, float size, Action<ISheetWriter> writeAction)
        {
            var settings = SettingsExstensions.FontSize(size);
            writer.WithTempSettings(settings, writeAction);
        }
    }
}