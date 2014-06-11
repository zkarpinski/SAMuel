Option Explicit
Dim FSO, InitFSO

'* create an instance of the File Browser
Set FSO = CreateObject("UserAccounts.CommonDialog")

'*setup the File Browser specifics
FSO.Filter = "All Files|*.*|Tiff|*.tiff|Adobe PDF|*.pdf|Microsoft Word|*.doc *docx"
FSO.Flags = &H0200 '*Allows multiple files to be selected
FSO.FilterIndex = 1
FSO.InitialDir = "c:\"

'* show the file browser and return the selection (or lack of) to InitFSO
InitFSO = FSO.ShowOpen

If InitFSO = False Then
    Wscript.Echo "Script Error: Please select a file!"
    Wscript.Quit
Else
    WScript.Echo "You selected the file: " & ObjFSO.FileName
End If
