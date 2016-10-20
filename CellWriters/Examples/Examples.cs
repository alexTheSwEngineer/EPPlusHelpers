using CellWriter.EPPlusHelpers.Excell;
using OfficeOpenXml;
using System.Drawing;
using System.IO;

namespace Examples
{
    class Examples
    {
        static void ISheetWriterExamples(string[] args)
        {
            using (var pck = new ExcelPackage(new FileInfo("D:\\ExampleExcellFile.xlsx")))
            {
                var sheet = new SheetWriter(pck.Workbook.Worksheets.Add("NewSheet"));
                
                sheet.Write("Cell1");        
                sheet.Write("Cell2"); 
                sheet.Write("Cell3", "Cell4");
                //  Will result with
                //     |  A  |  B  |  C  |  D  |
                //  |1||Cell1|Cell2|Cell3|Cell4|

                sheet.Write("Cell1", "Cell2", "Cell3", "Cell4");
                //  Will result with
                //     |  A  |  B  |  C  |  D  |
                //  |1||Cell1|Cell2|Cell3|Cell4|

                sheet.Write();
                sheet.Write("Cell2", "Cell3", "Cell4");
                //  Will result with
                //     |  A  |  B  |  C  |  D  |
                //  |1||     |Cell2|Cell3|Cell4|


                sheet.Write("Cell1", "Cell2");
                sheet.WriteLine();
                sheet.Write("Cell1", "Cell2");
                //  Will result with
                //     |  A  |  B  |
                //  |1||Cell1|Cell2|      
                //  |1||Cell1|Cell2|      

                sheet.WriteLine("Cell1", "Cell2");
                sheet.Write("Cell1", "Cell2");
                //  Will result with
                //     |  A  |  B  |
                //  |1||Cell1|Cell2|      
                //  |1||Cell1|Cell2|      



                sheet.WriteLine("Cell1", "Cell2");
                sheet.WriteLine();
                sheet.WriteLine();
                sheet.WriteLine("Cell1", "Cell2");
                //  Will result with
                //     |  A  |  B  |
                //  |1||Cell1|Cell2|      
                //  |2||     |     | 
                //  |3||     |     |  
                //  |4||Cell1|Cell2|  



                sheet.WithColor(Color.Red, x =>
                 {  
                     //x==sheet;
                     x.WriteLine("red", "red");
                     x.WriteLine("red");
                     x.Write();
                     x.Write("red");
                 });
                //  Will result with
                //     | A | B |
                //  |1||red|red|      
                //  |2||red|   | 
                //  |3||   |red|
                //All
                



            }
        }
    }
}
