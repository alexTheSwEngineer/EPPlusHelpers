# EPPlusHelpers
A small library that abstracts writing to excel sheets in a stream like manner.

#SheetWriter examples:

Instantiating:
```
var sheet = new SheetWriter(pck.Workbook.Worksheets.Add("NewSheet"));
```

#Using write()/writeLine():
```
	sheet.Write("Cell1", "Cell2", "Cell3", "Cell4");
	//is equivalent to:
	sheet.Write("Cell1");        
	sheet.Write("Cell2"); 
	sheet.Write("Cell3", "Cell4");	
	
     |  A  |  B  |  C  |  D  |
	 |1||Cell1|Cell2|Cell3|Cell4|
```
	 
	Any subsequent calls to Write() will continue to fill cells in the same row.



```
	sheet.Write();
	sheet.Write("Cell2", "Cell3", "Cell4");
	   |  A  |  B  |  C  |  D  |
	|1||     |Cell2|Cell3|Cell4|	
```
Even empty calls to Write() move the cursor to the next cell.


```
	sheet.Write("Cell1", "Cell2");
	sheet.WriteLine();
	sheet.Write("Cell1", "Cell2");	
    //is equivalent to:	
	sheet.WriteLine("Cell1", "Cell2");
	sheet.Write("Cell1", "Cell2");	
	
	   |  A  |  B  |
	|1||Cell1|Cell2|      
	|1||Cell1|Cell2|

```

```
	sheet.WriteLine("Cell1", "Cell2");
	sheet.WriteLine();
	sheet.WriteLine();
	sheet.WriteLine("Cell1", "Cell2");
	
		 |  A  |  B  |
	  |1||Cell1|Cell2|      
	  |2||     |     | 
	  |3||     |     |  
	  |4||Cell1|Cell2|  
```

#SheetWriter settings:

```
    sheet.WriteLine("none","none","none");
	sheet.WithColor(Color.Red, x =>
	 {  
		 //x==sheet;
		 x.WriteLine("red", "red");
		 x.WriteLine("red");
		 x.Write();
		 x.WriteLine("red");
	 });
	 sheet.WriteLine("none","none","none");
	 
	   |  A  |  B  |  C  |
    |1||none |none |none |         
	|2||red  |red  |     |      
	|3||red  |     |     |
	|4||     |red  |     |
    |5||none |none |none |   


All the cells that are have "red" wirten in them will have a red background, and the cells with "none" writen in them would have the default one.
What happens with the empty cells depends on the ShouldPadRows and ApplyModifiersToEmptyCells properties.
If ShouldPadRows was set to true all the fields [c1-c4] would be filled with cells with null value instead of being empty.
If both ApplyModifiersToEmptyCells and ShouldPadRows were set to true, every empty cell would have red background too.
If only ApplyModifiersToEmptyCells was set to true the only empty cell that has a red background would be A4.
```