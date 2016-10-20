using CellWriter.EPPlusHelpers.Excell;
using OfficeOpenXml;
using System.IO;

namespace Examples
{
    class Examples
    {
        static void Main(string[] args)
        {
            using (var pck = new ExcelPackage(new FileInfo("D:\\ExampleExcellFile.xlsx")))
            {
                var sheet = pck.Workbook.Worksheets["asd"];
                //var sheet = pck.Workbook.Worksheets.Add("asd");
                ISheetWriter writer = new SheetWriter(sheet, true,true);
                ISheetWriter writer2 = new SheetWriter(sheet, true);
                writer.Write();
                writer.Write("asd", "dsa");
                writer.WriteLine();
                writer.WriteLine();
                writer2.WriteLine();
                writer2.Write();
                writer2.Write();
                writer.Write("asd", "asd");
                writer.WriteLine("asd", "dsa");
                writer2.Write("asd","asd");
                pck.Save();
            }
        }
    }
}
