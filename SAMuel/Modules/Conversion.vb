Imports System.Drawing.Printing
Imports System.IO
Imports Ghostscript.NET
Imports ImageMagick
Imports Microsoft.Office.Interop.Word

Namespace Modules
    Module Conversion
        'Module Wide Varibles
        Private _sFileToPrint As String

        Friend Enum ConvertToType As Byte
            Tiff = 0
            Pdf = 1
        End Enum

        Sub ConvertImages(ByVal inputImage As String, ByVal outputFile As String, ByVal outType As ConvertToType)
            ''TODO Combine Combine ImgToPdf and ImgToTiff
        End Sub

        Sub ConvertDocuments(ByVal inputDoc As String, ByVal outputFile As String, ByVal outType As ConvertToType)
            ''TODO Combine Combine DocToPdf and DocToTiff
        End Sub

        Sub ImgToPdf(ByVal inputImg As String, ByVal outputPdf As String)
            ' Read image from file
            Using image As New MagickImage(inputImg)
                image.Write(outputPdf)
            End Using
        End Sub

        Sub ImgToTiff(inputFile As String, outputTiff As String)
            ' Read image from file
            Using image As New MagickImage(inputFile)
                image.CompressionMethod = CompressionMethod.Group4
                image.Write(outputTiff)
            End Using
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
            Dim gvi As GhostscriptVersionInfo = New GhostscriptVersionInfo(New System.Version(0, 0, 0), Directory.GetCurrentDirectory() + "\gsdll32.dll", String.Empty, GhostscriptLicense.GPL)
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
            'Set active printer to MODI
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

        ''' <summary>
        ''' Prints a word document to pdf using PDF995.
        ''' </summary>
        ''' <param name="inputDoc">Word document to be converted.</param>
        ''' <param name="outputPdf">Full path of the resulting pdf file.</param>
        ''' <param name="wordApp">Reference to an active word application.</param>
        ''' <remarks></remarks>
        Public Function DocToPdf(ByVal inputDoc As String, ByVal outputPdf As String, ByRef wordApp As Microsoft.Office.Interop.Word.Application) As Boolean
            Dim objWdDoc As Microsoft.Office.Interop.Word.Document

            'Initiate word application object and minimize it
            wordApp = CreateObject("Word.Application")
            wordApp.WindowState = Microsoft.Office.Interop.Word.WdWindowState.wdWindowStateMinimize

#If CONFIG = "Release" Then
            'Set active printer to PDF995
            Try
                wordApp.ActivePrinter = "PDF995"
            Catch ex As Exception
                MsgBox("Printer error. Is the 'PDF995' printer installed?", MsgBoxStyle.Critical)
                wordApp.Quit(False)
                Return False
            End Try

#End If

            objWdDoc = wordApp.Documents.Open(FileName:=inputDoc, ConfirmConversions:=False, AddToRecentFiles:=False)
            wordApp.Visible = False
            wordApp.WindowState = Microsoft.Office.Interop.Word.WdWindowState.wdWindowStateMinimize

            'Print to Pdf
            objWdDoc.PrintOut(Background:=False)
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