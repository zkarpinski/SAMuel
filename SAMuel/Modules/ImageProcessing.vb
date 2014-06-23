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
        myEncoderParameters = New EncoderParameters(2)

        ' Save the bitmap as a TIFF file with the following parameters
        myEncoderParameters.Param(0) = New EncoderParameter(cdEncoder, CType(24L, Int32))
        myEncoderParameters.Param(1) = New EncoderParameter(quEncoder, CType(25L, Int32))
        'myEncoderParameters.Param(2) = New EncoderParameter(cpEncoder, EncoderValue.CompressionLZW)
        myBitmap.Save(sSave, myImageCodecInfo, myEncoderParameters)

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

        'Determine the scale
        If wRatio < hRatio Then
            sRatio = wRatio
        Else : sRatio = hRatio
        End If

        width = bm.Width * sRatio
        height = bm.Height * sRatio
        myBitmap = New Bitmap(width, height)

        ' Create a newBitmap object based on a src
        myBitmap = New Bitmap(width, height)
        myBitmap.SetResolution(70, 70)

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
