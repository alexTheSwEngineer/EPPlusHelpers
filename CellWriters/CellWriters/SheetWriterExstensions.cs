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
                  .With(BgColor(color))
                  .Write(writeAction);
        }

        public static void WithSize(this ISheetWriter writer, float size, Action<ISheetWriter> writeAction)
        {
            writer.AsMonad()
                  .With(FontSize(size))
                  .Write(writeAction);
        }

        public static Action<ExcelRange> BgColor(Color color)
        {
            return cell =>
            {
                cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(color);
            };
        }

        public static Action<ExcelRange> FontSize(float size)
        {
            return cell =>
            {
                cell.Style.Font.Size = size;
            };
        }

    }
}