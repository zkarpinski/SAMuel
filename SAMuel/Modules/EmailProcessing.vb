Imports System.IO

Public Module EmailProcessing

    Function ValidateAttachmentType(ByVal sFile As String) As String
        Dim sFileExt As String
        sFileExt = Path.GetExtension(sFile).ToLower
        If sFileExt = ".tiff" Or sFileExt = ".png" Or _
                sFileExt = ".jpg" Or sFileExt = ".jpeg" Or _
                sFileExt = ".tif" Or sFileExt = ".gif" Or _
                sFileExt = ".bmp" Or sFileExt = ".pdf" Or _
                sFileExt = ".doc" Or sFileExt = ".docx" Then
            Return sFileExt
            'Known attachment types. 
        ElseIf sFileExt = ".psd" Or sFileExt = ".bin" Or _
                sFileExt = ".htm" Or sFileExt = ".html" Or _
                sFileExt = ".xps" Or sFileExt = ".txt." Then
            Return vbNullString
        Else
            LogAction(51, sFileExt)
            Return vbNullString
        End If
    End Function

End Module
