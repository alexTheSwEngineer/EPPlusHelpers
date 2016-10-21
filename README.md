# EPPlusHelpers
This is a small library that abstracts writing to excel sheets in a stream like manner. It is also usefull for 
managing formating and other cell settings.

#Usage:
```c#
using (var pckg = new ExcelPackage(new FileInfo("D:\\ExampleFile.xlsx")))
{
	var sheetWriter = new SheetWriter(pckg.Workbook.Worksheets.Add("Examples"));
	sheetWriter.SetUp()
			   .With(redFont);
			   .With(bigFont);
			   
	sheetWriter.WriteLine("with","big","red", "font")
			   .WithOverrideSettings(tempSetting1, x =>
			   {
				   x.WriteLine("some", "temp", "style","settings");
			   })
			   .WithAddedSettings(tempSetting2, x =>
			   {
				   x.WriteLine("some", "temp settings", "and","big","red","font");
			   })
			   .WithColor(Color.Blue,x=>
			   {
				  x.Write("cell1");
				  x.Write();//empty cell
				  x.Write("cell3");
				  x.WriteLine(); //go to next row
			   });
    pckg.Save();
}
```
Setting up a sheet writers default style is done through the SetUp()/ISettings api. 
Often there is need for some temporary change of the initial settings like in the case of creating table headers. This is most easilly done
by either of the methods: WithOverideSettings or WithAddedSettings.
The first temporaly uses only the settings from the parameter, and the later temporary combines the preexisting settings with the new ones sent as a parameter.
Any two setting can be combined :
```c#
 setting1.With(setting2);
```
Because conflicting settigs are possible, the last one added has priority. This can be cahnged by custom implementation of ISettings.
Examples:
```c#
var sheetWriter = new SheetWriter(pckg.Workbook.Worksheets.Add("Examples"));

ISettings redBackground = SettingsExstensions.BgColor(Color.Red);
ISettings bigFont = SettingsExstensions.FontSize(37);
ISettings mediumBorder = new Settings(cell => cell.Style.Border.BorderAround(ExcelBorderStyle.Medium));


sheetWriter.SetUp()
		   .With(redBackground);
		   
sheetWriter.WriteLine("with","red", "background")   //|1|
		   .WithOverideSettings(bigFont, x =>
		   {
			   x.WriteLine("with", "big", "font");  //|2|
		   })
		   .WithAddedSettings(bigFont, x =>
		   {
			   x.WriteLine("red", "big", "font");   //|3|
		   });

sheetWriter.SetUp().Clear();
sheetWriter.WriteLine("default", "formating");      //|4|

var combinedSettings = redBackground.With(bigFont);
sheetWriter.SetUp()
		   .With(combinedSettings);
sheetWriter.WriteLine("another", "big", "red", "text"); //|5|

sheetWriter.SetUp().Clear()
		   .With(redBackground)
		   .With(bigFont);
sheetWriter.WriteLine("big/red", "combined", "again");  //|6|

//This results in:
//   |   A   |   B     |   C      |   D   |
//|1||with   |red      |background|       |
//|2||WITH   |BIG      |FONT      |       |
//|3||RED    |BIG      |FONT      |       |
//|4||default|formating|          |       |
//|5||another|big      |red       |text   |
//|6||big/red|combined |again     |       |

```

#SheetWriter Properties:
Instantiating:
All of the SheetWriter properties are set up in the constructor. 
By default when a SheetWriter is instantiated the internal row/column pointers point to the end of the document.
If the need exists of multiple SheetWriters to simultaneously write to the same ExcelSheet, then the AllowMultipleWriterHanldes
must be set to true in each of the SheetWriters. Not doing so would result in errors. On the other hand, setting AllowMultipleWriterHanldes
to true when only one SheetWriter accesses the Worksheet will function correctly but will have a unneeded performance overhead. 
```c#
var sheet = new SheetWriter(pck.Workbook.Worksheets.Add("NewSheet"),AllowMultipleWriterHanldes:true);

```
The other properties have to do with what happens with empty cells:
```c#
	var redBackground = SettingsExstensions.BgColor(Color.Red)
	sheet.WriteLine("none","none","none");
		 .WithColor(Color.Red, x =>
		 {  
			 //x==sheet;
			 x.WriteLine("red", "red");
			 x.WriteLine("red");
			 x.Write();
			 x.WriteLine("red");
		 })
	     .WriteLine("none","none","none");
	 
//	   |  A  |  B  |  C  |
//  |1||none |none |none |         
//	|2||red  |red  |     |      
//	|3||red  |     |     |
//	|4||     |red  |     |
//  |5||none |none |none |   


//All the cells that are have "red" wirten in them will have a red background,
//and the cells with "none" writen in them would have the default one.
//What happens with the empty cells depends on the:
// ShouldPadRows and ApplyModifiersToEmptyCells properties.
//If ShouldPadRows was set to true:
//   all the fields [c1-c4] would be filled with cells with null value instead of being empty.
//If both ApplyModifiersToEmptyCells and ShouldPadRows were set to true:
//   every empty cell would have red background too.
//If only ApplyModifiersToEmptyCells was set to true:
//   the only empty cell that has a red background would be A4.
```

#Using write()/writeLine():

Calls to Write() will write to the cell that the internal row/column pointers point to, and it would move the column pointer to the next cell.
Any number of subsequent calls to Write() will continue to fill cells in the same row.
```c#
	sheet.Write("Cell1", "Cell2", "Cell3", "Cell4");
	//is equivalent to:
	sheet.Write("Cell1");        
	sheet.Write("Cell2"); 
	sheet.Write("Cell3", "Cell4");	
	
     |  A  |  B  |  C  |  D  |
	 |1||Cell1|Cell2|Cell3|Cell4|
```

Even empty calls to Write() move the cursor to the next cell.
```c#
	sheet.Write();
	sheet.Write("Cell2", "Cell3", "Cell4");
//	   |  A  |  B  |  C  |  D  |
//	|1||     |Cell2|Cell3|Cell4|	
```


Same goes with WriteLine(), ony this method moves the column pointer to the next cell for each parameter, and the row pointer only once per row.
Empty calls to WriteLine() will move the row pointer nonetheless. 
```c#
	sheet.Write("Cell1", "Cell2");
	sheet.WriteLine();
	sheet.Write("Cell1", "Cell2");	
    //is equivalent to:	
	sheet.WriteLine("Cell1", "Cell2");
	sheet.Write("Cell1", "Cell2");	
	
//	   |  A  |  B  |
//	|1||Cell1|Cell2|      
//	|1||Cell1|Cell2|

```
```c#
	sheet.WriteLine("Cell1", "Cell2");
	sheet.WriteLine();
	sheet.WriteLine();
	sheet.WriteLine("Cell1", "Cell2");
	
//		 |  A  |  B  |
//	  |1||Cell1|Cell2|      
//	  |2||     |     | 
//	  |3||     |     |  
//	  |4||Cell1|Cell2|  
```

