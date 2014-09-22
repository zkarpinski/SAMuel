Imports System.IO
Imports Outlook = Microsoft.Office.Interop.Outlook

Public Module EmailProcessing

    ''' <summary>
    ''' Creates a data object for each email within the given outlook folder.
    ''' </summary>
    ''' <param name="olFolder">Outlook folder to work out of.</param>
    ''' <returns>List of SAM_EMAIL Objects.</returns>
    ''' <remarks></remarks>
    Function GetEmails(ByRef olFolder As Outlook.MAPIFolder) As List(Of SAM_Email)
        Dim sEmail As SAM_Email
        Dim emailItem As Outlook.MailItem
        Dim samEmails As New List(Of SAM_Email)
        For Each emailItem In olFolder.Items
            Try
                sEmail = New SAM_Email(emailItem)
                samEmails.Add(sEmail)
            Catch ex As Exception
                LogAction(action:="An email was skipped because " & ex.Message)
                'Move to next email
                Continue For
            End Try
        Next
        Return samEmails
    End Function


    Function ValidateAttachmentType(ByVal sFile As String) As String
        Dim sFileExt As String
        sFileExt = Path.GetExtension(sFile).ToLower
        If sFileExt = ".tiff" Or sFileExt = ".png" Or _
                sFileExt = ".jpg" Or sFileExt = ".jpeg" Or _
                sFileExt = ".tif" Or sFileExt = ".gif" Or _
                sFileExt = ".bmp" Or sFileExt = ".pdf" Or _
                sFileExt = ".doc" Or sFileExt = ".docx" Then
            Return sFileExt
            'Known attachment types that are skipped (for now).     
        ElseIf sFileExt = ".psd" Or sFileExt = ".bin" Or _
                sFileExt = ".htm" Or sFileExt = ".html" Or _
                sFileExt = ".xps" Or sFileExt = ".txt" Or _
                sFileExt = ".msg" Or sFileExt = ".vcf" Then
            Return vbNullString
        Else
            LogAction(51, sFileExt)
            Return vbNullString
        End If
    End Function

    Public Sub DeleteSavedAttachments(ByVal sDirectory As String)
        If (Directory.Exists(sDirectory)) Then Directory.Delete(sDirectory, True)
    End Sub



End Module
