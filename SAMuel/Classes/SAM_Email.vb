Public Class SAM_Email

    Public Property Subject As String
    Public Property From As String
    Public Property Body As String
    Public Property Account As String
    Public Property Attachments As New List(Of String)
    Public Property IsValid As Boolean = True
    Public Property IsBillAccount As Boolean = True

    Public Sub New(ByRef objEmail As Object)
        Dim sFile As String
        Dim sPath As String
        Dim rand As New Random

        Subject = objEmail.Subject
        From = objEmail.SenderName
        Body = objEmail.body

        sPath = My.Settings.savePath + "emails\" + From + "\"
        GlobalModule.CheckFolder(sPath)

        If objEmail.Attachments.Count > 0 Then
            For Each value In objEmail.Attachments
                'Added random number for same filename handling cases
                sFile = sPath & rand.Next(10000).ToString & value.FileName
                value.SaveAsFile(sFile)
                Attachments.Add(sFile)
            Next
        Else
            IsValid = False
        End If
    End Sub
    Public Sub Regex()
        Dim strRegex As String
        strRegex = GlobalModule.RegexAccount(Subject)
        If strRegex = "ACC# NOT FOUND" Then
            strRegex = GlobalModule.RegexCustomer(Subject)
            If strRegex = "CUST# NOT FOUND" Then
                strRegex = GlobalModule.RegexAccount(Body)
                If strRegex = "ACC# NOT FOUND" Then
                    strRegex = GlobalModule.RegexCustomer(Body)
                    If strRegex = "CUST# NOT FOUND" Then
                        strRegex = "UNKNOWN"
                        IsValid = False
                    End If
                End If
            End If
        End If

        Account = strRegex
    End Sub

    ReadOnly Property AttachmentsCount As Integer
        Get
            Return Attachments.Count
        End Get
    End Property

End Class
