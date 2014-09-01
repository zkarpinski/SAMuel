Imports Outlook = Microsoft.Office.Interop.Outlook
Public Class SAM_Email
    Implements IDisposable
    Public Property Subject As String
    Public Property From As String
    Public Property SenderEmailAddress As String
    Public Property Body As String
    Public Property Account As String
    Public Property Attachments As New List(Of String)
    Public Property IsValid As Boolean = True
    Public Property IsBillAccount As Boolean = True
    Public Property EmailObject As Object

    Private disposed As Boolean = False ' To detect dispose redundant calls

    Public Sub New(ByRef objEmail As Object)
        Me.Subject = objEmail.Subject
        Me.From = objEmail.SenderName
        Me.Body = objEmail.body
        Me.SenderEmailAddress = objEmail.SenderEmailAddress
        Me.Account = Me.Regex()
        Me.EmailObject = objEmail
    End Sub

    Private Function Regex() As String
        Dim strAccNum As String = "X"
        Dim strRegFormats() As String = {"\d{5}-\d{5}", "\d{10}", "\d{9}"} ' 'Regex Formats for Bill Account Numbers and Customer Numbers.

        'Use regular expressions to search for an account number.
        For Each rFormat As String In strRegFormats
            strAccNum = GlobalModule.RegexAcc(Me.Subject, rFormat)
            If strAccNum <> "X" Then Exit For
            strAccNum = GlobalModule.RegexAcc(Me.Body, rFormat)
            If strAccNum <> "X" Then Exit For
        Next
        'If nothing was found, flag the email.
        If strAccNum = "X" Then
            strAccNum = ""
            Me.IsValid = False
        End If

        Return strAccNum
    End Function

    Public Sub DownloadAttachments()
        Dim sFile As String
        Dim sPath As String
        Dim rand As New Random
        Dim subFolder As String

        subFolder = SenderEmailAddress + "\"
        ' TODO Remove ALL invalid path characters from string.
        sPath = My.Settings.savePath + "emails\" + subFolder
        GlobalModule.CheckFolder(sPath)

        If Me.EmailObject.Attachments.Count > 0 Then
            For Each value In Me.EmailObject.Attachments
                'Added random number for same filename handling cases
                sFile = sPath + value.FileName + "_" + rand.Next(10000).ToString
                value.SaveAsFile(sFile)
                Me.Attachments.Add(sFile)
            Next
        Else
            IsValid = False
        End If
    End Sub

    ReadOnly Property AttachmentsCount As Integer
        Get
            Return Me.Attachments.Count
        End Get
    End Property

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposed Then
            If disposing Then
                ' Free other state (managed objects).
            End If
            ' Free your own state (unmanaged objects).
            Me.EmailObject = Nothing
            Me.Attachments = Nothing
            Me.Subject = vbNullString
            Me.Body = vbNullString
        End If
        Me.disposed = True
    End Sub

#Region " IDisposable Support "
    ' This code added by Visual Basic to     ' correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.        
        ' Put cleanup code in
        ' Dispose(ByVal disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overrides Sub Finalize()
        ' Do not change this code.
        ' Put cleanup code in
        ' Dispose(ByVal disposing As Boolean) above.
        Dispose(False)
        MyBase.Finalize()
    End Sub
#End Region

End Class
