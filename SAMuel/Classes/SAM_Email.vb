Public Class SAM_Email

    Private mSubject As String
    Private mFrom As String
    Private mBody As String
    Private mCustAcc As String
    Private mAttachments As New List(Of String)

    Private IsVaidAcc As Boolean = True

    Public Sub New(ByRef objEmail As Object)
        Dim sFile As String
        Dim sPath As String = My.Settings.savePath

        mSubject = objEmail.Subject
        mFrom = objEmail.SenderName
        mBody = objEmail.body

        If objEmail.Attachments.Count > 0 Then
            For Each value In objEmail.Attachments
                sFile = sPath + value.FileName
                value.SaveAsFile(sFile)
                mAttachments.Add(sFile)
            Next
        End If
    End Sub

    ReadOnly Property AttachmentsCount As Integer
        Get
            Return mAttachments.Count
        End Get
    End Property

    ReadOnly Property Subject As String
        Get
            Return mSubject
        End Get
    End Property

    ReadOnly Property From As String
        Get
            Return mFrom
        End Get
    End Property

    ReadOnly Property Account As String
        Get
            Return mCustAcc
        End Get
    End Property
End Class
