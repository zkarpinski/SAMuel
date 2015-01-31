Imports Microsoft.Office.Interop.Outlook
Imports SAMuel.Classes.SamEmail
Imports SAMuel.Modules.GlobalModule

Namespace Classes
    Public Class EmailProcessor
        Private _oApp As Microsoft.Office.Interop.Outlook.Application = New Microsoft.Office.Interop.Outlook.Application
        Private _wApp As Microsoft.Office.Interop.Word.Application = Nothing
        Private _workingOutlookFolder As MAPIFolder
        Private _saveFolder As String
        Private _samEmails As List(Of SamEmail)

        Public Property CompletedEmailCount As String
        Public Property StartTime As DateTime
        Public Property EndTime As DateTime
        Public Property TotalTime As TimeSpan
        Public Function EmailCount() As Integer
            Return _samEmails.Count
        End Function

        Public Sub New(sFolder As String)
            _oApp = New Microsoft.Office.Interop.Outlook.Application
            _wApp = New Microsoft.Office.Interop.Word.Application()
            _saveFolder = sFolder
        End Sub

        'User chooses the folder to work from.
        Public Function ChooseOutlookFolder() As Boolean
            _WorkingOutlookFolder = _oApp.GetNamespace("MAPI").PickFolder()
            If _WorkingOutlookFolder Is Nothing Then
                Return False
            Else
                Return True
            End If
        End Function

        Public Sub GetEmails()
            Me.StartTime = DateTime.Now
            Dim sEmail As SamEmail
            Dim obj As Object
            Dim emailItem As Microsoft.Office.Interop.Outlook.MailItem
            _samEmails = New List(Of SamEmail)
            For Each obj In _workingOutlookFolder.Items
                Try
                    'Check object type.
                    'http://mztools.com/articles/2006/MZ2006013.aspx
                    If Microsoft.VisualBasic.Information.TypeName(obj) = "MailItem" Then
                        emailItem = CType(obj, Microsoft.Office.Interop.Outlook.MailItem)
                        sEmail = New SamEmail(emailItem)
                        _samEmails.Add(sEmail)
                    End If

                Catch ex As System.Exception
                    '   LogAction(action:=" An email was skipped because " & ex.Message)
                    'Move to next email
                    Continue For
                End Try
            Next
        End Sub

        Public Sub ProcessEmails()
            For Each sEmail In _samEmails
                sEmail.DownloadAttachments(ATT_FOLDER)
                If sEmail.AttachmentCount > 0 Then

                End If
            Next
        End Sub

        Private Function ProcessAttachments()

        End Function



    End Class
End Namespace
