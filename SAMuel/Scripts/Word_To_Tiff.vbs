Const msoFileDialogOpen = 1

Set objWord = CreateObject("Word.Application")
activePrinter =  objWord.ActivePrinter
'*objWord.ActivePrinter = "Microsoft Office Document Image Writer"
objWord.ActivePrinter = "Microsoft XPS Document Writer"
objWord.ChangeFileOpenDirectory("C:\")

objWord.FileDialog(msoFileDialogOpen).Title = "Select the files to be printed"
objWord.FileDialog(msoFileDialogOpen).AllowMultiSelect = True

If objWord.FileDialog(msoFileDialogOpen).Show = -1 Then
    objWord.WindowState = 2
    Set objFiles = objWord.FileDialog(msoFileDialogOpen).SelectedItems
    For Each objFile in objFiles
       Set openFile = objWord.Documents.Open(objFile)
       openFile.PrintOut()
       Set openFile = nothing
    Next   
End If

objWord.ActivePrinter = activePrinter
objWord.Quit

