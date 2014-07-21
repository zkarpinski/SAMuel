Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Word = Microsoft.Office.Interop.Word
Imports System.IO

Module ConvertToTiff
    Sub WordDocs(sFiles() As String)
        'Converts Word Documents to .tif using MODI
        Dim objWdDoc As Word.Document
        Dim objWord As Word.Application
        Dim sDestination As String = My.Settings.savePath + "converted\"
        Dim sFileName As String

        'Initiate word application object and minimize it
        objWord = CreateObject("Word.Application")
        objWord.WindowState = Word.WdWindowState.wdWindowStateMinimize
        'Set active printer to Fax
        objWord.ActivePrinter = "Microsoft Office Document Image Writer"

        'Verify output folder exists
        GlobalModule.CheckFolder(sDestination)

        For Each value In sFiles
            'Get the file name
            sFileName = Path.GetFileNameWithoutExtension(value)
            'Open the document within word and don't prompt  for conversion.
            objWdDoc = objWord.Documents.Open(FileName:=value, ConfirmConversions:=False, AddToRecentFiles:=False, [ReadOnly]:=True, Visible:=False)
            objWord.Visible = False
            'Print to Tiff
            objWdDoc.PrintOut(PrintToFile:=True, OutputFileName:=sDestination & sFileName & ".tif")
            'Release document
            objWdDoc.Close()
            'Progress the progress bar
            'frmMain.ProgressBar.Value += 1
        Next
        'Cleanup
        objWord.Quit()
        objWdDoc = Nothing
        objWord = Nothing
    End Sub
End Module
