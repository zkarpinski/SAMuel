Imports System.Drawing.Imaging
Imports System.IO
Imports System.Runtime.InteropServices

Module ImageProcessing

    ''' <summary>
    ''' Compresses, encodes and saves the image as a tiff
    ''' </summary>
    ''' <param name="myBitmap">Bitmap ByRef to be saved.</param>
    ''' <param name="sSave">Directory and file name to save file as</param>
    ''' <remarks>Reference http://msdn.microsoft.com/en-US/en-en/library/bb882583.aspx </remarks>
    Sub CompressTiff(ByRef myBitmap As Bitmap, sSave As String)
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
    End Sub

    ''' <summary>
    ''' Resizes the image
    ''' </summary>
    ''' <param name="bm">
    ''' Bitmap ByRef to be resized
    ''' </param>
    ''' <returns> MemoryStream of resized image</returns>
    Public Function ResizeImage(ByRef bm As Bitmap) As Bitmap
        Dim width As Integer
        Dim height As Integer
        Dim myBitmap As Bitmap
        Dim wRatio As Double = 1700 / bm.Width
        Dim hRatio As Double = 2200 / bm.Height
        Dim sRatio As Double

        'Rotate image
        If bm.Width > bm.Height Then
            bm.RotateFlip(RotateFlipType.Rotate90FlipNone)
            wRatio = 1700 / bm.Width
            hRatio = 2200 / bm.Height
        End If


        'If bm.Height > 2200 Or bm.Width > 1700 Then

        '    'Determine the scale
        '    If wRatio > hRatio Then
        '        sRatio = wRatio
        '    Else : sRatio = hRatio
        '    End If

        '    width = bm.Width * sRatio
        '    height = bm.Height * sRatio
        'Else
        '    width = bm.Width
        '    height = bm.Height
        'End If

        width = bm.Width
        height = bm.Height

        ' Create a newBitmap object based on a src
        myBitmap = New Bitmap(width, height)
        myBitmap.SetResolution(200, 200)

        Dim g As Graphics = Graphics.FromImage(myBitmap)
        g.InterpolationMode = Drawing2D.InterpolationMode.Default
        g.DrawImage(bm, New Rectangle(0, 0, width, height), New Rectangle(0, 0, bm.Width, bm.Height), GraphicsUnit.Pixel)


        'Release resources
        g.Dispose()
        g = Nothing
        bm.Dispose()
        bm = Nothing

        'Return the bitstream
        Return myBitmap
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

    Public Function ConvertToRGB(original As Bitmap) As Bitmap
        Dim g As Graphics
        Dim newImage As Bitmap = New Bitmap(original.Width, original.Height, PixelFormat.Format32bppArgb)
        newImage.SetResolution(original.HorizontalResolution, original.VerticalResolution)
        g = Graphics.FromImage(newImage)
        g.DrawImageUnscaled(original, 0, 0)
        g.Dispose()
        g = Nothing
        Return newImage
    End Function

    Public Function ConvertToBitonal(original As Bitmap) As Bitmap
        Dim source As Bitmap
        ' If original bitmap is not already in 32 BPP, ARGB format, then convert
        If (original.PixelFormat <> PixelFormat.Format32bppArgb) Then

            source = New Bitmap(original.Width, original.Height, PixelFormat.Format32bppArgb)
            source.SetResolution(original.HorizontalResolution, original.VerticalResolution)
            Dim g As Graphics = Graphics.FromImage(source)
            g.DrawImageUnscaled(original, 0, 0)
            g.Dispose()
            g = Nothing
        Else
            source = original
        End If

        '// Lock source bitmap in memory
        Dim sourceData As BitmapData = source.LockBits(New Rectangle(0, 0, source.Width, source.Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb)

        '// Copy image data to binary array
        Dim imageSize As Integer = sourceData.Stride * sourceData.Height
        Dim sourceBuffer(imageSize) As Byte
        Marshal.Copy(sourceData.Scan0, sourceBuffer, 0, imageSize)

        ' Unlock source bitmap
        source.UnlockBits(sourceData)

        ' Create destination bitmap
        Dim destination As Bitmap = New Bitmap(source.Width, source.Height, PixelFormat.Format1bppIndexed)
        destination.SetResolution(original.HorizontalResolution, original.VerticalResolution)

        ' Lock destination bitmap in memory
        Dim destinationData As BitmapData = destination.LockBits(New Rectangle(0, 0, destination.Width, destination.Height), ImageLockMode.WriteOnly, PixelFormat.Format1bppIndexed)

        ' Create destination buffer
        imageSize = destinationData.Stride * destinationData.Height
        Dim destinationBuffer(imageSize) As Byte

        Dim sourceIndex As Integer = 0
        Dim destinationIndex As Integer = 0
        Dim pixelTotal As Integer = 0
        Dim destinationValue As Byte = 0
        Dim pixelValue As Integer = 128
        Dim height As Integer = source.Height
        Dim width As Integer = source.Width
        Dim threshold As Integer = 500

        '// Iterate lines
        For y As Integer = 0 To height - 1 Step 1

            sourceIndex = y * sourceData.Stride
            destinationIndex = y * destinationData.Stride
            destinationValue = 0
            pixelValue = 128
            '// Iterate pixels
            For x As Integer = 0 To width Step 1
                ' Compute pixel brightness (i.e. total of Red, Green, and Blue values) - Thanks murx
                '                           B                             G                              R
                pixelTotal = sourceBuffer(sourceIndex)
                pixelTotal += sourceBuffer(sourceIndex + 1)
                pixelTotal += sourceBuffer(sourceIndex + 2)
                If (pixelTotal > threshold) Then destinationValue += Convert.ToByte(pixelValue)
                If (pixelValue = 1) Then

                    destinationBuffer(destinationIndex) = destinationValue
                    destinationIndex += 1
                    destinationValue = 0
                    pixelValue = 128

                Else
                    pixelValue >>= 1
                    sourceIndex += 4
                End If
            Next
            If (pixelValue <> 128) Then destinationBuffer(destinationIndex) = destinationValue

        Next

        'Copy binary image data to destination bitmap
        Marshal.Copy(destinationBuffer, 0, destinationData.Scan0, imageSize)

        ' Unlock destination bitmap
        destination.UnlockBits(destinationData)

        '// Dispose of source if not originally supplied bitmap
        'If (source  original) Then
        source.Dispose()


        ' Return
        Return destination
    End Function


End Module
