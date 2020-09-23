<div align="center">

## Simple ole drag and drop


</div>

### Description

This demonstrates a simple drag and drop of a file from windows explorer into a text box then grab the filename and path (this is part of the code im using to make a map install program for quake2)
 
### More Info
 
a dropped file from windows explorer

you must create a textbox called text1 and make sure its big enough so its not too fiddly to drop a file onto. also make sure for text1 oledragmode and oledrop mode in the properties are set to manual

incorrect file type

none known


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Stewart MacFarlane](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/stewart-macfarlane.md)
**Level**          |Beginner
**User Rating**    |4.6 (51 globes from 11 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/stewart-macfarlane-simple-ole-drag-and-drop__1-6598/archive/master.zip)





### Source Code

```
Private Sub text1_OLEDragDrop(Data As DataObject, Effect As Long _
, Button As Integer, Shift As Integer, X As Single, Y As Single)
' Prepare a variable (numfiles) and pass the number of files
' dropped onto text1 to this variable
Dim numFiles As Integer
 numFiles = Data.Files.Count
' an example how to trap 1 file (can be modified to trap as many
' or as little amount by changing the > 1 to > {new value}) then
' display a message box telling user the maximum allowed file drops)
' then exit the sub
If numFiles > 1 Then
	MsgBox "Only allows 1 file at a time in beta version! Sorry!"_
	,vbOKOnly, "Ooops beta version"
	Exit Sub
end if
' check the attributes of the file being dropped and see if it is a
' directory, if it is then warn user that only files are valid to drop
' and exit the sub
If (GetAttr(Data.Files(1))) = vbDirectory Then
 MsgBox "Sorry this beta version only allows files not directories to be installed"
 Exit Sub
End If
' check the file is the correct file type (using its extension)
' if not then warn user and exit the sub
If LCase(Right(Data.Files(1), 3)) <> LCase("bsp") Then
 MsgBox "This file is not a quake 2 map (*.bsp)"
 Exit Sub
End If
' tell user the drag and drop was succesful
MsgBox Data.Files(1) + " installed"
' code here to install file
' or do what ever you need
' data.files(1) is a string holding the path and filename of the dropped file
' using a for..next loop you can control multiple files dropped at once
' replacing the 1 with the for..next variable and using numfiles to find out
' the maximum for..next value
End Sub
```

