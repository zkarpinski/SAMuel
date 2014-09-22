Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Word = Microsoft.Office.Interop.Word
Imports System.IO
Imports System.Drawing.Printing

Imports Ghostscript.NET


Module Conversion

    Dim sFileToPrint As String

    Sub imagesToTiff(sFiles() As String)

    End Sub

    Sub imgToTiff(inputFile As String, outputTiff As String)

        Dim printDocument As New PrintDocument()
        AddHandler printDocument.PrintPage, AddressOf printDocument_PrintPage

        'Controller hides the "printing page x of y dialog (http://stackoverflow.com/questions/5511138/can-i-disable-the-printing-page-x-of-y-dialog)
        Dim printController As PrintController = New StandardPrintController
        printDocument.PrintController = printController

        'Set printer settings
        printDocument.PrinterSettings.PrintToFile = True
#If CONFIG = "Release" Then
        printDocument.PrinterSettings.PrinterName = "Microsoft Office Document Image Writer"
#End If
        If Not printDocument.PrinterSettings.IsValid Then
            MsgBox("Printer error. Is the 'Microsoft Office Document Image Writer' printer installed?", MsgBoxStyle.Critical)
            Return
        End If

        sFileToPrint = inputFile

        'Print the image
        printDocument.PrinterSettings.PrintFileName = outputTiff
        printDocument.Print()
    End Sub

    ''' <summary>
    ''' Converts PDFs to multi-page tiffs
    ''' </summary>
    ''' <param name="inputPDF"></param>
    ''' <param name="outputTiff"></param>
    ''' <remarks></remarks>
    Public Sub pdfToTiff(inputPDF As String, outputTiff As String)

        'Run ghostscript with arguments.
        'http://www.ghostscript.com/doc/current/Use.htm
        Dim Args() As String = {"-q", "-dNOPAUSE", "-dBATCH", "-dSAFER", "-sDEVICE=tiffg4", "-sPAPERSIZE=letter", _
                                "-dNumRenderingThreads=" & Environment.ProcessorCount.ToString(), _
                                "-sOutputFile=" & outputTiff, "-f" & inputPDF}
        Dim gvi As GhostscriptVersionInfo = New GhostscriptVersionInfo(New Version(0, 0, 0), Directory.GetCurrentDirectory() + "\gsdll32.dll", String.Empty, GhostscriptLicense.GPL)
        Dim processor As Processor.GhostscriptProcessor = New Processor.GhostscriptProcessor(gvi, True)

        processor.StartProcessing(Args, Nothing)
    End Sub

    ''' <summary>
    ''' Prints a word document to tiff using Microsoft Office Document Image Writer
    ''' </summary>
    ''' <param name="inputDoc">Word document to be converted.</param>
    ''' <param name="outputTiff">Full path of the resulting tiff file.</param>
    ''' <param name="wordApp">Reference to an active word application.</param>
    ''' <remarks></remarks>
    Public Function docToTiff(ByVal inputDoc As String, ByVal outputTiff As String, ByRef wordApp As Word.Application) As Boolean
        Dim objWdDoc As Word.Document

        'Initiate word application object and minimize it
        wordApp = CreateObject("Word.Application")
        wordApp.WindowState = Word.WdWindowState.wdWindowStateMinimize

#If CONFIG = "Release" Then
        'Set active printer to Fax
        Try
            wordApp.ActivePrinter = "Microsoft Office Document Image Writer"
        Catch ex As Exception
            MsgBox("Printer error. Is the 'Microsoft Office Document Image Writer' printer installed?", MsgBoxStyle.Critical)
            wordApp.Quit(False)
            Return False
        End Try

#End If

        objWdDoc = wordApp.Documents.Open(FileName:=inputDoc, ConfirmConversions:=False, AddToRecentFiles:=False)
        wordApp.Visible = False
        wordApp.WindowState = Word.WdWindowState.wdWindowStateMinimize

        'Print to Tiff
        objWdDoc.PrintOut(PrintToFile:=True, OutputFileName:=outputTiff, Background:=False)
        Threading.Thread.Sleep(1000)
        'Release document and close word.
        objWdDoc.Close()
        wordApp.Quit()
        Return True
    End Function

    Private Sub printDocument_PrintPage(ByVal sender As Object, ByVal e As PrintPageEventArgs)
        'Resize rotate image if needed then print within page bounds.
        Dim mBitmap As Bitmap = Bitmap.FromFile(sFileToPrint)
        mBitmap = ImageProcessing.ResizeImage(mBitmap)
        e.Graphics.DrawImage(CType(mBitmap, Image), 0, 0, e.PageBounds.Width, e.PageBounds.Height)
    End Sub

End Module
