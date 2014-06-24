Imports System.IO
Imports System.Drawing.Imaging

Public Module EmailProcessing

    ''' <summary>
    ''' Adds a string watermark to the passed image
    ''' </summary>
    ''' <param name="img">Image ByRef to add watermark to</param>
    ''' <param name="accountNumber">String to be added as watermark</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Add_Watermark(ByRef img As Image, ByVal accountNumber As String) As Image
        'Adds the account number as a watermark to the image
        'http://bytescout.com/products/developer/watermarkingsdk/how_to_add_text_watermark_to_image_using_dotnet_in_vb.html
        Dim graphics As Graphics
        Dim font As Font
        Dim brush As SolidBrush
        Dim rBrush As SolidBrush
        Dim point As PointF

        ' Account Number Watermark Settings
        font = New Font("Times New Roman", 20.0F)
        point = New PointF(50, 50) '(x,y)
        brush = New SolidBrush(Color.FromArgb(200, Color.Black))
        rBrush = New SolidBrush(Color.FromArgb(160, Color.White))

        ' Create graphics from image
        graphics = Drawing.Graphics.FromImage(img)
        ' Draw the Account Number Watermark
        graphics.FillRectangle(rBrush, point.X, point.Y, 230, 40)
        graphics.DrawString(accountNumber, font, brush, point)

        'Clear Resources
        graphics.Dispose()
        graphics = Nothing
        Return img
    End Function

    ''' <summary>
    ''' Changes the image to grayscale using ColorMatrix
    ''' </summary>
    ''' <param name="img">Image ByRef to be changed to gray scale</param>
    ''' <remarks></remarks>
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
        g = Nothing

    End Sub

    Function ValidateAttachmentType(ByVal sFile As String) As String
        Dim sFileExt As String
        sFileExt = Path.GetExtension(sFile).ToLower
        If sFileExt = ".tiff" Or sFileExt = ".png" Or _
                sFileExt = ".jpg" Or sFileExt = ".jpeg" Or _
                sFileExt = ".tif" Or sFileExt = ".gif" Or _
                sFileExt = ".bmp" Or sFileExt = ".pdf" Then
            Return sFileExt
        Else
            Return vbNullString
        End If
    End Function


    'Test Function
    Sub ParsePDFImgs(pdfPath As String) 'As List(Of pdflib.clsPDFImage)
        Dim count As Integer
        Dim imagelist As List(Of pdflib.clsPDFImage)
        Dim image As pdflib.clsPDFImage
        Dim page As pdflib.clsPDFPage

        Dim myPDF As New pdflib.ClsPDF(pdfPath)

        'We get the images of the pdf
        imagelist = myPDF.GetImages

        'We save them
        count = 0
        For Each image In imagelist
            count += 1
            image.Save(My.Settings.savePath + "_image_" + count.ToString)
        Next

        'Get Images from page 1 only
        page = myPDF.GetPage(1)
        imagelist = page.GetImages

        myPDF.Close()

    End Sub

End Module
