Imports System.IO
Imports System.Drawing.Imaging

Public Module EmailProcessing
    Public Function Add_Watermark(ByVal img As Image, ByVal accountNumber As String, Optional ByVal suffix As String = vbNullString) As String
        'Adds the account number as a watermark to the image
        'http://bytescout.com/products/developer/watermarkingsdk/how_to_add_text_watermark_to_image_using_dotnet_in_vb.html
        Dim graphics As Graphics
        Dim font As Font
        Dim brush As SolidBrush
        Dim point As PointF
        Dim savedFile As String

        'Determine if file needs a suffix
        If suffix = vbNullString Then
            savedFile = My.Settings.savePath + accountNumber + ".jpg"
        Else
            savedFile = My.Settings.savePath + accountNumber + "_" + suffix + ".jpg"
        End If

        ' Account Number Watermark Settings
        font = New Font("Times New Roman", 30.0F)
        point = New PointF(500, 50) '(x,y)
        brush = New SolidBrush(Color.FromArgb(150, Color.Black))

        ' Create graphics from image
        graphics = Drawing.Graphics.FromImage(img)
        ' Draw the Account Number Watermark
        graphics.DrawString(accountNumber, font, brush, point)
        ' Save image
        img.Save(savedFile)
        'Clear Resources
        graphics.Dispose()
        graphics = Nothing
        img.Dispose()
        img = Nothing
        Return savedFile
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
                                       "testImage")

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
End Module
