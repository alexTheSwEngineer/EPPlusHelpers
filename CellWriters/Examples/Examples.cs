using CellWriters.Core;
using CellWriters.Exstensions;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.IO;

namespace Examples
{
    class Examples
    {
        static void Main(string[] args)
        {
            using (var pckg = new ExcelPackage(new FileInfo("D:\\ExampleFile.xlsx")))
            {
                var sheetWriter = new SheetWriter(pckg.Workbook.Worksheets.Add("Examples"));

                ISettings redBackground = SettingsExstensions.BgColor(Color.Red);
                ISettings bigFont = SettingsExstensions.FontSize(37);
                ISettings mediumBorder = new Settings(cell => cell.Style.Border.BorderAround(ExcelBorderStyle.Medium));
                

                sheetWriter.SetUp()
                           .With(redBackground);
                           
                sheetWriter.WriteLine("with","red", "background")
                           .WithOverrideSettings(bigFont, x =>
                           {
                               x.WriteLine("with", "big", "font");
                           })
                           .WithAddedSettings(bigFont, x =>
                           {
                               x.WriteLine("red", "big", "font");
                           });

                sheetWriter.SetUp().Clear();
                sheetWriter.WriteLine("default", "formating");

                var combinedSettings = redBackground.With(bigFont);
                sheetWriter.SetUp()
                           .With(combinedSettings);
                sheetWriter.WriteLine("another", "big", "red", "text");

                sheetWriter.SetUp().Clear()
                           .With(redBackground)
                           .With(bigFont);
                sheetWriter.WriteLine("big/red", "combined", "again");

                pckg.Save();
            }
        }
    }
}
