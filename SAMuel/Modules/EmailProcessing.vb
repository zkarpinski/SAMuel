Imports System.IO
Imports System.Drawing.Imaging

Public Module EmailProcessing

    ''' <summary>
    ''' Adds a string watermark to the passed image
    ''' </summary>
    ''' <param name="myBitmap">Bitmap ByRef to add watermark to</param>
    ''' <param name="accountNumber">String to be added as watermark</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Add_Watermark(ByRef myBitmap As Bitmap, ByVal accountNumber As String) As Image
        'Adds the account number as a watermark to the image
        'http://bytescout.com/products/developer/watermarkingsdk/how_to_add_text_watermark_to_image_using_dotnet_in_vb.html
        Dim graphics As Graphics
        Dim font As Font
        Dim brush As SolidBrush
        Dim rBrush As SolidBrush
        Dim point As PointF

        ' Account Number Watermark Settings
        font = New Font(My.Settings.wmFont, 9.0F)
        point = New PointF(5, 0) '(x,y)
        brush = New SolidBrush(Color.FromArgb(250, Color.Black))
        rBrush = New SolidBrush(Color.FromArgb(250, Color.White))

        ' Create graphics from image
        graphics = Drawing.Graphics.FromImage(myBitmap)
        ' Draw the Account Number Watermark
        graphics.FillRectangle(rBrush, point.X, point.Y + 1, 128, 20)
        graphics.DrawString(accountNumber, font, brush, point)

        'Clear Resources
        graphics.Dispose()
        graphics = Nothing
        Return myBitmap
    End Function

    ''' <summary>
    ''' Changes the image to grayscale using ColorMatrix
    ''' </summary>
    ''' <param name="myBitmap">Bitmap ByRef to be changed to gray scale</param>
    ''' <remarks></remarks>
    Sub MakeGrayscale(ByRef myBitmap As Bitmap)
        Dim g As Graphics = Graphics.FromImage(myBitmap)
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
        g.DrawImage(myBitmap, New Rectangle(0, 0, myBitmap.Width, myBitmap.Height), 0, 0, myBitmap.Width, myBitmap.Height, GraphicsUnit.Pixel, attributes)
        'dispose the Graphics object
        g.Dispose()
        g = Nothing

    End Sub

    Function ValidateAttachmentType(ByVal sFile As String) As String
        Dim sFileExt As String
        sFileExt = Path.GetExtension(sFile).ToLower
        If sFileExt = ".tiff" Or sFileExt = ".png" Or _
                sFileExt = ".jpg" Or sFileExt = ".jpeg" Or _
                sFileExt = ".tif" Or sFileExt = ".gif" Or _
                sFileExt = ".bmp" Or sFileExt = ".pdf" Or _
                sFileExt = ".doc" Or sFileExt = ".docx" Then
            Return sFileExt
        Else
            Return vbNullString
        End If
    End Function
End Module
