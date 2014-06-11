''Convert Word to Tif files
''Created by: Zachary Karpinski
''Purpose: Expediate converting files from word to tif.

''Requirements:
''  Microsoft Office 2003+
''  Microsoft Office Document Imaging(MODI) OR Microsoft XPS Document Writer

'Constants
Const msoFileDialogOpen = 1

'Create a Word Application Object
Set objWord = CreateObject("Word.Application")
'Store the current default printer
activePrinter =  objWord.ActivePrinter

'Change printer to MODI/XPS
''''objWord.ActivePrinter = "Microsoft Office Document Image Writer"
objWord.ActivePrinter = "Microsoft XPS Document Writer"

'Microsoft Office File Dialog Settings
objWord.ChangeFileOpenDirectory("C:\")
objWord.FileDialog(msoFileDialogOpen).Title = "Select the files to be converted."
objWord.FileDialog(msoFileDialogOpen).AllowMultiSelect = True

'When user clicks open button..
If objWord.FileDialog(msoFileDialogOpen).Show = -1 Then
    objWord.WindowState = 2 
    Set objFiles = objWord.FileDialog(msoFileDialogOpen).SelectedItems
    For Each objFile in objFiles
        'Open each document but don't show    
        Set openFile = objWord.Documents.Open(objFile, Visible:=False)
        'Print
        openFile.PrintOut()
        'Wait for system to catchup then release object
        WScript.Sleep 300
        openFile.Close()
        Set openFile = nothing
    Next   
End If

'Restore inital default printer
objWord.ActivePrinter = activePrinter

'Close Word and release
objWord.Quit
Set objWord = nothing

