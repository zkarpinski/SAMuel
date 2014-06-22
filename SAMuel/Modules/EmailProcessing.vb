Imports System.IO
Imports System.Drawing.Imaging

Public Module EmailProcessing
    Public Function Add_Watermark(ByRef img As Image, ByVal accountNumber As String) As Image
        'Adds the account number as a watermark to the image
        'http://bytescout.com/products/developer/watermarkingsdk/how_to_add_text_watermark_to_image_using_dotnet_in_vb.html
        Dim graphics As Graphics
        Dim font As Font
        Dim brush As SolidBrush
        Dim rBrush As SolidBrush
        Dim point As PointF

        ' Account Number Watermark Settings
        font = New Font("Times New Roman", 30.0F)
        point = New PointF(50, 50) '(x,y)
        brush = New SolidBrush(Color.FromArgb(150, Color.Red))
        rBrush = New SolidBrush(Color.FromArgb(100, Color.White))

        ' Create graphics from image
        graphics = Drawing.Graphics.FromImage(img)
        ' Draw the Account Number Watermark
        graphics.FillRectangle(rBrush, point.X, point.Y, 250, 40)
        graphics.DrawString(accountNumber, font, brush, point)

        'Clear Resources
        graphics.Dispose()
        graphics = Nothing
        Return img
    End Function

    ''**NOTE**This function needs an overhaul
    Public Sub Convert_Image_To_Tif(fileNames As String, isMultiPage As Boolean)
        'Converts the image to a tiff
        'http://code.msdn.microsoft.com/windowsdesktop/VBTiffImageConverter-f8fdfd7f/sourcecode?fileId=26346&pathId=565080007
        Using encoderParams As New EncoderParameters(1)
            Dim tiffCodecInfo As ImageCodecInfo = ImageCodecInfo.GetImageEncoders(). _
                   First(Function(ie) ie.MimeType = "image/tiff")

            Dim tiffPaths As String() = Nothing
            If isMultiPage Then
                tiffPaths = New String(1) {}
                Dim tiffImg As Image = Nothing
                Try
                    Dim i As Integer = 0
                    While i < fileNames.Length
                        If i = 0 Then
                            tiffPaths(i) = [String].Format("{0}\{1}.tiff",
                                        Path.GetDirectoryName(fileNames(i)),
                                        Path.GetFileNameWithoutExtension(fileNames(i)))

                            ' Initialize the first frame of multi-page tiff. 
                            tiffImg = Image.FromFile(fileNames(i))
                            encoderParams.Param(0) = New EncoderParameter(
                                Encoder.SaveFlag, DirectCast(EncoderValue.MultiFrame, 
                                EncoderParameterValueType))
                            tiffImg.Save(tiffPaths(i), tiffCodecInfo, encoderParams)
                        Else
                            ' Add additional frames. 
                            encoderParams.Param(0) = New EncoderParameter(
                                Encoder.SaveFlag,
                                DirectCast(EncoderValue.FrameDimensionPage, 
                                EncoderParameterValueType))
                            Using frame As Image = Image.FromFile(fileNames(i))
                                tiffImg.SaveAdd(frame, encoderParams)
                            End Using
                        End If

                        If i = fileNames.Length - 1 Then
                            ' When it is the last frame, flush the resources and closing. 
                            encoderParams.Param(0) = New EncoderParameter(
                                Encoder.SaveFlag,
                                DirectCast(EncoderValue.Flush, EncoderParameterValueType))
                            tiffImg.SaveAdd(encoderParams)
                        End If
                        System.Math.Max(System.Threading.Interlocked.Increment(i), i - 1)
                    End While
                Finally
                    If tiffImg IsNot Nothing Then
                        tiffImg.Dispose()
                        tiffImg = Nothing
                    End If
                End Try
            Else
                tiffPaths = New String(fileNames.Length) {}

                Dim i As Integer = 0
                While i < fileNames.Length
                    tiffPaths(i) = [String].Format("{0}{1}.tiff",
                                       My.Settings.savePath,
                                       "test")

                    ' Save as individual tiff files. 
                    Using tiffImg As Image = Image.FromFile(fileNames)
                        tiffImg.Save(tiffPaths(i), ImageFormat.Tiff)
                    End Using
                    System.Math.Max(System.Threading.Interlocked.Increment(i), i - 1)
                End While
            End If
        End Using
    End Sub

    Private Sub Convert_Image_To_Tif(sEditedImg As String, Optional p2 As Object = Nothing, Optional p3 As Boolean = Nothing)
        Throw New NotImplementedException
    End Sub

    Sub Resize_Image(ByRef img As Image)
        'following code resizes picture to fit

        Dim bm As New Bitmap(img)
        Dim width As Integer
        Dim height As Integer
        Dim newBM As Bitmap
        Dim wRatio As Double = 1700 / bm.Width
        Dim hRatio As Double = 2200 / bm.Height
        Dim sRatio As Double

        'Determine the scale
        If wRatio < hRatio Then
            sRatio = wRatio
        Else : sRatio = hRatio
        End If

        width = bm.Width * sRatio
        height = bm.Height * sRatio
        newBM = New Bitmap(width, height)

        Dim g As Graphics = Graphics.FromImage(newBM)
        g.InterpolationMode = Drawing2D.InterpolationMode.Default
        g.DrawImage(bm, New Rectangle(0, 0, width, height), New Rectangle(0, 0, bm.Width, bm.Height), GraphicsUnit.Pixel)

        g.Dispose()
        bm.Dispose()

        'image path.
        newBM.Save(My.Settings.savePath + "test12123.jpg")

        newBM.Dispose()

    End Sub

    Sub MakeGrayscale(ByRef img As Image)
        Dim g As Graphics = Graphics.FromImage(img)
        Dim attributes As ImageAttributes = New ImageAttributes()
        'The grayscale ColorMatrix http://msdn.microsoft.com/en-us/library/system.drawing.imaging.colormatrix(v=vs.110).aspx
        Dim colorMatrixElements As Single()() = { _
            New Single() {0.3F, 0.3F, 0.3F, 0, 0}, _
            New Single() {0.59F, 0.59F, 0.59F, 0, 0}, _
            New Single() {0.11F, 0.11F, 0.11F, 0, 0}, _
            New Single() {0, 0, 0, 1, 0}, _
            New Single() {0, 0, 0, 0, 1} _
            }
        Dim colorMatrix As New ColorMatrix(colorMatrixElements)

        'set the color matrix attribute
        attributes.SetColorMatrix(colorMatrix)

        'draw the original image on the new image
        'using the grayscale color matrix
        g.DrawImage(img, New Rectangle(0, 0, img.Width, img.Height), 0, 0, img.Width, img.Height, GraphicsUnit.Pixel, attributes)
        'dispose the Graphics object
        g.Dispose()

    End Sub

    Function ValidateAttachmentType(ByVal sFile As String) As String
        Dim sFileExt As String
        sFileExt = Path.GetExtension(sFile).ToLower
        If sFileExt = ".tiff" Or sFileExt = ".png" Or _
                sFileExt = ".jpg" Or sFileExt = ".jpeg" Or _
                sFileExt = ".tif" Or sFileExt = ".gif" Or _
                sFileExt = ".bmp" Then
            Return sFileExt
        Else
            Return vbNullString
        End If
    End Function

End Module
