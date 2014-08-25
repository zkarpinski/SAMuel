Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Word = Microsoft.Office.Interop.Word
Imports System.IO
Imports System.Drawing.Printing

Module ConvertToTiff

    Dim sFileToPrint As String

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
        Try
            objWord.ActivePrinter = "Microsoft Office Document Image Writer"
        Catch ex As Exception
            MsgBox("Printer error. Is the 'Microsoft Office Document Image Writer' printer installed?", MsgBoxStyle.Critical)
            objWord.Quit()
            Exit Sub
        End Try

        'Verify output folder exists
        GlobalModule.CheckFolder(sDestination)

        'Update UI
        frmMain.ProgressBar.Maximum = sFiles.Length
        frmMain.ProgressBar.Value = 0

        For Each value In sFiles
            frmMain.lblStatus.Text = String.Format("Converting {0} of {1}", frmMain.ProgressBar.Value + 1, sFiles.Length)
            frmMain.Refresh()
            'Get the file name
            sFileName = Path.GetFileNameWithoutExtension(value)
            'Open the document within word and don't prompt  for conversion.
            objWdDoc = objWord.Documents.Open(FileName:=value, ConfirmConversions:=False)
            objWord.Visible = False
            objWord.WindowState = Word.WdWindowState.wdWindowStateMinimize
            'Print to Tiff
            objWdDoc.PrintOut(PrintToFile:=True, OutputFileName:=sDestination & sFileName & ".tif")
            'Release document
            objWdDoc.Close()
            'Progress the progress bar
            frmMain.ProgressBar.Value += 1
        Next

        objWord.Quit()
    End Sub

    Sub ImageFiles(sFiles() As String)
        Dim sDestination As String = My.Settings.savePath + "converted\"
        Dim sFileName As String
        Dim sOutTIFF As String

        Dim printDocument As New PrintDocument()
        AddHandler printDocument.PrintPage, AddressOf printDocument_PrintPage

        'Controller hides the "printing page x of y dialog (http://stackoverflow.com/questions/5511138/can-i-disable-the-printing-page-x-of-y-dialog)
        Dim printController As PrintController = New StandardPrintController
        printDocument.PrintController = printController

        'Set printer settings
        printDocument.PrinterSettings.PrintToFile = True
        printDocument.PrinterSettings.PrinterName = "Microsoft Office Document Image Writer"
        If Not printDocument.PrinterSettings.IsValid Then
            MsgBox("Printer error. Is the 'Microsoft Office Document Image Writer' printer installed?", MsgBoxStyle.Critical)
            Exit Sub
        End If

        'Verify output folder exists
        GlobalModule.CheckFolder(sDestination)

        frmMain.ProgressBar.Maximum = sFiles.Length
        frmMain.ProgressBar.Value = 0

        For Each value In sFiles
            frmMain.lblStatus.Text = String.Format("Converting {0} of {1}", frmMain.ProgressBar.Value + 1, sFiles.Length)
            frmMain.Refresh()
            sFileToPrint = value
            sFileName = Path.GetFileNameWithoutExtension(value)
            sOutTIFF = [String].Format("{0}{1}.tiff", sDestination, sFileName)

            'Print the image
            printDocument.PrinterSettings.PrintFileName = sOutTIFF
            printDocument.Print()

            frmMain.ProgressBar.Value += 1
        Next
    End Sub

    Private Sub printDocument_PrintPage(ByVal sender As Object, ByVal e As PrintPageEventArgs)
        'Resize rotate image if needed then print within page bounds.
        Dim mBitmap As Bitmap = Bitmap.FromFile(sFileToPrint)
        mBitmap = ImageProcessing.ResizeImage(mBitmap)
        e.Graphics.DrawImage(CType(mBitmap, Image), 0, 0, e.PageBounds.Width, e.PageBounds.Height)
    End Sub

End Module
