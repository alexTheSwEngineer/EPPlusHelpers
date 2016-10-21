using CellWriter.EPPlusHelpers.Excell;
using CellWriters.Core;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CellWriters.Exstensions
{
    public  class SettingsExstensions
    {
        public static ISettings BgColor(Color color)
        {
            return new Settings( cell =>
            {
                cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(color);
            });
        }

        public static ISettings FontSize(float size)
        {
            return new Settings(cell =>
            {
                cell.Style.Font.Size = size;
            });
        }
    }
}
