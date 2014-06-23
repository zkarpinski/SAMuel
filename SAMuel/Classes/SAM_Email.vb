Public Class SAM_Email

    Private mSubject As String
    Private mFrom As String
    Private mBody As String
    Private mCustAcc As String
    Private mAttachments As New List(Of String)

    Private _IsValid As Boolean = True

    Public Sub New(ByRef objEmail As Object)
        Dim sFile As String
        Dim sPath As String

        mSubject = objEmail.Subject
        mFrom = objEmail.SenderName
        mBody = objEmail.body

        sPath = My.Settings.savePath + "emails\" + mFrom + "\"
        GlobalModule.CheckFolder(sPath)
        If objEmail.Attachments.Count > 0 Then
            For Each value In objEmail.Attachments
                sFile = sPath + value.FileName
                value.SaveAsFile(sFile)
                mAttachments.Add(sFile)
            Next
        Else
            _IsValid = False
        End If
    End Sub
    Public Sub Regex()
        Dim strRegex As String
        strRegex = GlobalModule.RegexAccount(mSubject)
        If strRegex = "ACC# NOT FOUND" Then
            strRegex = GlobalModule.RegexCustomer(mSubject)
            If strRegex = "CUST# NOT FOUND" Then
                strRegex = GlobalModule.RegexAccount(mBody)
                If strRegex = "ACC# NOT FOUND" Then
                    strRegex = GlobalModule.RegexCustomer(mBody)
                    If strRegex = "CUST# NOT FOUND" Then
                        strRegex = "UNKNOWN"
                        _IsValid = False
                    End If
                End If
            End If
        End If

        mCustAcc = strRegex
    End Sub

    Public Property Account() As String
        Get
            Return mCustAcc
        End Get
        Set(ByVal value As String)
            mCustAcc = value
        End Set
    End Property

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

    ReadOnly Property Attachments As List(Of String)
        Get
            Return mAttachments
        End Get
    End Property
    ReadOnly Property IsValid As Boolean
        Get
            Return _IsValid
        End Get
    End Property

End Class
