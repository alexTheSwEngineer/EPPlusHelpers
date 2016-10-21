using CellWriter.EPPlusHelpers.Excell;
using CellWriters.Exstensions;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
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
                var redBackground = SettingsExstensions.BgColor(Color.Red);
                var bigFont = SettingsExstensions.FontSize(37);
                var sheetWriter = new SheetWriter(pckg.Workbook.Worksheets.Add("Examples"));

                sheetWriter.SetUp()
                           .With(redBackground);
                           
                sheetWriter.WriteLine("with","red", "background")
                           .WithTempSettings(bigFont, x =>
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
