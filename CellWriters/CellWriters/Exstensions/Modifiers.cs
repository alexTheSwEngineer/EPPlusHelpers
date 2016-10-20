using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CellWriters.Exstensions
{
    public class Modifiers
    {
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
