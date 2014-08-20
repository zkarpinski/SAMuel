Imports Outlook = Microsoft.Office.Interop.Outlook
Public Class SAM_Email
    Implements IDisposable
    Public Property Subject As String
    Public Property From As String
    Public Property Body As String
    Public Property Account As String
    Public Property Attachments As New List(Of String)
    Public Property IsValid As Boolean = True
    Public Property IsBillAccount As Boolean = True
    Public Property EmailObject As Object

    Private disposed As Boolean = False ' To detect dispose redundant calls

    Public Sub New(ByRef objEmail As Object)
        Subject = objEmail.Subject
        From = objEmail.SenderName
        Body = objEmail.body
        Regex()
        EmailObject = objEmail

    End Sub

    Private Sub Regex()
        Dim strRegex As String
        strRegex = GlobalModule.RegexAccount(Subject)
        If strRegex = "ACC# NOT FOUND" Then
            strRegex = GlobalModule.RegexCustomer(Subject)
            If strRegex = "CUST# NOT FOUND" Then
                strRegex = GlobalModule.RegexAccount(Body)
                If strRegex = "ACC# NOT FOUND" Then
                    strRegex = GlobalModule.RegexCustomer(Body)
                    If strRegex = "CUST# NOT FOUND" Then
                        strRegex = ""
                        IsValid = False
                    End If
                End If
            End If
        End If

        Account = strRegex
    End Sub

    Public Sub DownloadAttachments()
        Dim sFile As String
        Dim sPath As String
        Dim rand As New Random
        Dim subFolder As String

        subFolder = From + "\"
        subFolder = subFolder.Replace(":", "")
        subFolder = subFolder.Replace("*", "")
        subFolder = subFolder.Replace("?", "")
        subFolder = subFolder.Replace("<", "")
        subFolder = subFolder.Replace(">", "")
        ' TODO Remove invalid path characters
        sPath = My.Settings.savePath + "emails\" + subFolder
        GlobalModule.CheckFolder(sPath)

        If EmailObject.Attachments.Count > 0 Then
            For Each value In EmailObject.Attachments
                'Added random number for same filename handling cases
                sFile = sPath & rand.Next(10000).ToString & value.FileName
                value.SaveAsFile(sFile)
                Attachments.Add(sFile)
            Next
        Else
            IsValid = False
        End If
    End Sub

    ReadOnly Property AttachmentsCount As Integer
        Get
            Return Attachments.Count
        End Get
    End Property

    ' IDisposable
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If Not Me.disposed Then
            If disposing Then
                ' Free other state (managed objects).
            End If
            ' Free your own state (unmanaged objects).
            EmailObject = Nothing
            Attachments = Nothing
            Subject = vbNullString
            Body = vbNullString
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
