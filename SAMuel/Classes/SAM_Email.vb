Imports System.Text.RegularExpressions
Imports SAMuel.Modules

Namespace Classes
	Public Class SamEmail
		Implements IDisposable
		Public Property Subject As String
		Public Property From As String
		Public Property SenderEmailAddress As String
		Public Property Body As String
		Public Property Accounts As Array
		Public Property Attachments As New List(Of SAM_File)
		Public Property IsValid As Boolean = True
		Public Property IsBillAccount As Boolean = True
		Public Property EmailObject As Object

		Private _disposed As Boolean = False ' To detect dispose redundant calls

        Public Sub New(ByRef objEmail As Object)
			Me.Subject = objEmail.Subject
			Me.From = objEmail.SenderName
			Me.Body = objEmail.body
			Me.SenderEmailAddress = objEmail.SenderEmailAddress
			Me.Accounts = Me.emRegex()
			Me.EmailObject = objEmail
		End Sub


		Private Function emRegex() As Array
			'Regex Formats for Bill Account Numbers and Customer Numbers.
			Dim strRegFormats() As String = {"((\D|^)(\d{5}-\d{5})(?!\d))", "((\D|^)(\d{10})(?!\d))", "((\D|^)(\d{9})(?!\d))"}
			Dim searchStr As String
			Dim rReg As Regex

			searchStr = Me.Subject & " " & Me.Body
            'Use regular expressions to search for an account number.
            For Each rFormat As String In strRegFormats
				Dim arr = rReg.Matches(searchStr, rFormat).OfType(Of Match)().[Select](Function(m) m.Groups(3).Value).ToArray
				If arr.Count > 0 Then
					Return arr
				End If
			Next

			'If nothing was found, flag the email.
			Me.IsValid = False

			Return New String() {}
		End Function

		Public Sub DownloadAttachments(ByVal sPath As String)
			Dim sFile As String
			Dim rand As New Random
			Dim subFolder As String

			subFolder = SenderEmailAddress + "\"
            ' TODO Remove ALL invalid path characters from string.
            sPath = sPath + subFolder
			GlobalModule.CheckFolder(sPath)

			If Me.EmailObject.Attachments.Count > 0 Then
				For Each value In Me.EmailObject.Attachments
                    'Added random number for same filename handling cases
                    sFile = sPath + rand.Next(100).ToString + "_" + value.FileName
					value.SaveAsFile(sFile)
					Me.Attachments.Add(New SAM_File(sFile))
				Next
			Else
				IsValid = False
			End If
		End Sub

		ReadOnly Property AttachmentCount As Integer
			Get
				Return Me.Attachments.Count
			End Get
		End Property

        ' IDisposable
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
			If Not Me._disposed Then
				If disposing Then
                    ' Free other state (managed objects).
                End If
                ' Free your own state (unmanaged objects).
                Me.EmailObject = Nothing
				Me.Attachments = Nothing
				Me.Subject = vbNullString
				Me.Body = vbNullString
			End If
			Me._disposed = True
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
End Namespace