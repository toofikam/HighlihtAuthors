Really don't understand what I am doing.

Seem to be like having a 50 tons stone to crush a mosquitos.

SO sorry for what you will find here, it is really to understand what we are doing in that univers and most of all what for!!??


Edit...
Totally impossible to understand anything to what is that f... for. Impossible to create another file!!!
To busy to spend sometime to try to understand something so baldly build...

Below is the macro (unfinished but working) for LIbre Office :

sub UnderlineAuthors
rem ----------------------------------------------------------------------
rem define variables
dim document   as object
dim dispatcher as object

rem Add the number of auhtor to highlight
dim authors(2) as string
dim i As Integer

rem Fill the list of author to be highlight...
authors(0)="TOOFIK AM"
authors(1)="NEXT ONE"
rem etc...



rem ----------------------------------------------------------------------
rem get access to the document
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")



dim args1(18) as new com.sun.star.beans.PropertyValue
dim args2(0) as new com.sun.star.beans.PropertyValue
args2(0).Name = "Underline.LineStyle"
args2(0).Value = 1



For i = 0 To 37

args1(0).Name = "SearchItem.StyleFamily"
args1(0).Value = 2
args1(1).Name = "SearchItem.CellType"
args1(1).Value = 0
args1(2).Name = "SearchItem.RowDirection"
args1(2).Value = true
args1(3).Name = "SearchItem.AllTables"
args1(3).Value = false
args1(4).Name = "SearchItem.Backward"
args1(4).Value = false
args1(5).Name = "SearchItem.Pattern"
args1(5).Value = false
args1(6).Name = "SearchItem.Content"
args1(6).Value = false
args1(7).Name = "SearchItem.AsianOptions"
args1(7).Value = false
args1(8).Name = "SearchItem.AlgorithmType"
args1(8).Value = 0
args1(9).Name = "SearchItem.SearchFlags"
args1(9).Value = 65536
args1(10).Name = "SearchItem.SearchString"
args1(10).Value = authors(i)
args1(12).Name = "SearchItem.Locale"
args1(12).Value = 255
args1(13).Name = "SearchItem.ChangedChars"
args1(13).Value = 2
args1(14).Name = "SearchItem.DeletedChars"
args1(14).Value = 2
args1(15).Name = "SearchItem.InsertedChars"
args1(15).Value = 2
args1(16).Name = "SearchItem.TransliterateFlags"
args1(16).Value = 1280
args1(17).Name = "SearchItem.Command"
args1(17).Value = 1
args1(18).Name = "Quiet"
args1(18).Value = true

dispatcher.executeDispatch(document, ".uno:ExecuteSearch", "", 0, args1())
dispatcher.executeDispatch(document, ".uno:Underline", "", 0, args2())

Next i

end sub


