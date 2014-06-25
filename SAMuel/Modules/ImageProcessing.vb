Imports System.Drawing.Imaging
Imports System.IO

Module ImageProcessing

    ''' <summary>
    ''' Compresses, encodes and saves the image as a tiff
    ''' </summary>
    ''' <param name="img">Image ByRef to be saved.</param>
    ''' <param name="sSave">Directory and file name to save file as</param>
    ''' <remarks>Reference http://msdn.microsoft.com/en-US/en-en/library/bb882583.aspx </remarks>
    Sub CompressTiff(ByRef img As Image, sSave As String)
        Dim myBitmap As New Bitmap(img)
        Dim myImageCodecInfo As ImageCodecInfo
        Dim cdEncoder As Encoder
        Dim cpEncoder As Encoder
        Dim quEncoder As Encoder
        Dim myEncoderParameters As EncoderParameters

        ' Get an ImageCodecInfo object that represents the TIFF codec.
        myImageCodecInfo = GetEncoder(ImageFormat.Tiff)

        ' Create an Encoder object based on the GUID 
        cpEncoder = Encoder.Compression
        cdEncoder = Encoder.ColorDepth
        quEncoder = Encoder.Quality

        ' Create an EncoderParameters object. 
        myEncoderParameters = New EncoderParameters(3)

        ' Save the bitmap as a TIFF file with the following parameters
        myEncoderParameters.Param(0) = New EncoderParameter(cdEncoder, CType(My.Settings.colorDepth, Int32))
        myEncoderParameters.Param(1) = New EncoderParameter(quEncoder, CType(80L, Int32))

        'Choose compression type from settings
        Select Case My.Settings.tiffCompression
            Case "None"
                myEncoderParameters.Param(2) = New EncoderParameter(cpEncoder, EncoderValue.CompressionNone)
            Case "CCITT3"
                myEncoderParameters.Param(2) = New EncoderParameter(cpEncoder, EncoderValue.CompressionCCITT3)
            Case "CCITT4"
                myEncoderParameters.Param(2) = New EncoderParameter(cpEncoder, EncoderValue.CompressionCCITT4)
            Case "LZW"
                myEncoderParameters.Param(2) = New EncoderParameter(cpEncoder, EncoderValue.CompressionLZW)
            Case "RLE"
                myEncoderParameters.Param(2) = New EncoderParameter(cpEncoder, EncoderValue.CompressionRle)
            Case Else
                myEncoderParameters.Param(2) = New EncoderParameter(cpEncoder, EncoderValue.CompressionNone)
        End Select


        Try
            myBitmap.Save(sSave, myImageCodecInfo, myEncoderParameters)
        Catch ex As Exception
            LogAction(99, ex.ToString)
            MsgBox("Compression error. Try different settings in options.")
            myBitmap.Dispose()
            myBitmap = Nothing
            Stop
        End Try

        myBitmap.Dispose()
        myBitmap = Nothing
    End Sub

    ''' <summary>
    ''' Resizes the image
    ''' </summary>
    ''' <param name="img">
    ''' Image ByRef to be resized
    ''' </param>
    ''' <returns> MemoryStream of resized image</returns>
    Public Function ResizeImage(ByRef img As Image) As MemoryStream
        Dim bm As New Bitmap(img)
        Dim width As Integer
        Dim height As Integer
        Dim myBitmap As Bitmap
        Dim wRatio As Double = 1700 / bm.Width
        Dim hRatio As Double = 2200 / bm.Height
        Dim sRatio As Double
        Dim ms As New MemoryStream

        If bm.Height > 2200 Or bm.Width > 1700 Then

            'Determine the scale
            If wRatio < hRatio Then
                sRatio = wRatio
            Else : sRatio = hRatio
            End If

            width = bm.Width * sRatio
            height = bm.Height * sRatio
        Else
            width = bm.Width
            height = bm.Height
        End If

        ' Create a newBitmap object based on a src
        myBitmap = New Bitmap(width, height)
        myBitmap.SetResolution(96, 96)

        Dim g As Graphics = Graphics.FromImage(myBitmap)
        g.InterpolationMode = Drawing2D.InterpolationMode.Default
        g.DrawImage(bm, New Rectangle(0, 0, width, height), New Rectangle(0, 0, bm.Width, bm.Height), GraphicsUnit.Pixel)

        myBitmap.Save(ms, ImageFormat.Jpeg)

        'Release resources
        g.Dispose()
        g = Nothing
        bm.Dispose()
        bm = Nothing
        myBitmap.Dispose()
        myBitmap = Nothing

        'Return the bitstream
        Return ms
    End Function
    Private Function GetEncoder(ByVal format As ImageFormat) As ImageCodecInfo

        Dim codecs As ImageCodecInfo() = ImageCodecInfo.GetImageDecoders()

        Dim codec As ImageCodecInfo
        For Each codec In codecs
            If codec.FormatID = format.Guid Then
                Return codec
            End If
        Next codec
        Return Nothing

    End Function


End Module
