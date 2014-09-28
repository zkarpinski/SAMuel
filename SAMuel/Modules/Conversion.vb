Imports System.Drawing.Printing
Imports System.IO
Imports Ghostscript.NET

Namespace Modules
    Module Conversion
        'Module Wide Varibles
        Private _sFileToPrint As String
        Public Const pdf995ini As String = "C:\Program Files\pdf995\res\pdf995.ini"

        Sub SetupWordPdfConv(ByRef wordApplication As Object)
            wordApplication.ActivePrinter = "PDF995"
            'Copy old PDF995 INF
            File.Copy(pdf995ini, SAVE_FOLDER + "original_pdf995.ini", True)
            'Replace with new PDF995 inf
            Return
        End Sub

        Sub CreatePdf995Ini(ByVal outputFolder As String, ByVal documentName As String)
            'Removes the ending /
            outputFolder = outputFolder.Remove(outputFolder.Length - 1)

            'Creates or Overwrites the XML file
            Dim fs1 As FileStream = New FileStream(pdf995ini, FileMode.Create, FileAccess.Write)
            Dim s1 As StreamWriter = New StreamWriter(fs1)

            'Create ini file
            s1.Write("[Parameters]" & vbCrLf)
            s1.Write("Display Readme=0" & vbCrLf)
            s1.Write("AutoLaunch=0" & vbCrLf) 'Disable PDF from opening after print.
            s1.Write("Quiet=1" & vbCrLf) 'Disables UI?
            s1.WriteLine("Use GPL Ghostscript=1") '?
            s1.Write("Document Name=" & documentName & vbCrLf) 'Source file
            s1.Write("Initial Dir=" & outputFolder & vbCrLf) 'Output folder?
            s1.Write("Output File=SAMEASDOCUMENT" & vbCrLf) 'Output file name
            s1.WriteLine("Output Folder=" & outputFolder) 'Output Folder

            s1.Write("[OmniFormat]" & vbCrLf)
            s1.Write("Accept EULA=1" & vbCrLf)

            s1.Close()
            fs1.Close()
        End Sub

        Sub ImgToPdf(ByVal inputImg As String, ByVal outputPdf As String)

            Dim printDocument As New PrintDocument()
            AddHandler printDocument.PrintPage, AddressOf printDocument_PrintPage

            'Controller hides the "printing page x of y dialog (http://stackoverflow.com/questions/5511138/can-i-disable-the-printing-page-x-of-y-dialog)
            Dim printController As PrintController = New StandardPrintController
            printDocument.PrintController = printController

            'Set printer settings
#If CONFIG = "Release" Then
            printDocument.PrinterSettings.PrinterName = "PDF995"
#End If
            If Not printDocument.PrinterSettings.IsValid Then
                MsgBox("Printer error. Is the 'PDF995' printer installed?", MsgBoxStyle.Critical)
                Return
            End If

            _sFileToPrint = inputImg

            'Print the image
            'printDocument.PrinterSettings.PrintFileName = outputPdf
            printDocument.DocumentName = outputPdf
            printDocument.Print()

        End Sub

        Sub ImgToTiff(inputFile As String, outputTiff As String)

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

            _sFileToPrint = inputFile

            'Print the image
            printDocument.PrinterSettings.PrintFileName = outputTiff
            printDocument.Print()
        End Sub

        ''' <summary>
        ''' Converts PDFs to multi-page tiffs
        ''' </summary>
        ''' <param name="inputPdf"></param>
        ''' <param name="outputTiff"></param>
        ''' <remarks></remarks>
        Public Sub PdfToTiff(inputPdf As String, outputTiff As String)

            'Run ghostscript with arguments.
            'http://www.ghostscript.com/doc/current/Use.htm
            Dim args() As String = {"-q", "-dNOPAUSE", "-dBATCH", "-dSAFER", "-sDEVICE=tiffg4", "-sPAPERSIZE=letter", _
                                    "-dNumRenderingThreads=" & Environment.ProcessorCount.ToString(), _
                                    "-sOutputFile=" & outputTiff, "-f" & inputPdf}
            Dim gvi As GhostscriptVersionInfo = New GhostscriptVersionInfo(New Version(0, 0, 0), Directory.GetCurrentDirectory() + "\gsdll32.dll", String.Empty, GhostscriptLicense.GPL)
            Dim processor As Processor.GhostscriptProcessor = New Processor.GhostscriptProcessor(gvi, True)

            processor.StartProcessing(args, Nothing)
        End Sub

        ''' <summary>
        ''' Prints a word document to tiff using Microsoft Office Document Image Writer
        ''' </summary>
        ''' <param name="inputDoc">Word document to be converted.</param>
        ''' <param name="outputTiff">Full path of the resulting tiff file.</param>
        ''' <param name="wordApp">Reference to an active word application.</param>
        ''' <remarks></remarks>
        Public Function DocToTiff(ByVal inputDoc As String, ByVal outputTiff As String, ByRef wordApp As Microsoft.Office.Interop.Word.Application) As Boolean
            Dim objWdDoc As Microsoft.Office.Interop.Word.Document

            'Initiate word application object and minimize it
            wordApp = CreateObject("Word.Application")
            wordApp.WindowState = Microsoft.Office.Interop.Word.WdWindowState.wdWindowStateMinimize

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
            wordApp.WindowState = Microsoft.Office.Interop.Word.WdWindowState.wdWindowStateMinimize

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
            Dim mBitmap As Bitmap = Bitmap.FromFile(_sFileToPrint)
            mBitmap = ImageProcessing.ResizeImage(mBitmap)
            e.Graphics.DrawImage(CType(mBitmap, Image), 0, 0, e.PageBounds.Width, e.PageBounds.Height)
        End Sub

    End Module
End Namespace