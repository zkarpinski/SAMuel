Public Class frmEmails

    Private Sub frmEmails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Center Form
        Dim mainScreen As Screen = Screen.FromPoint(Me.Location)
        Dim X As Integer = (mainScreen.WorkingArea.Width - Me.Width)
        Dim Y As Integer = mainScreen.WorkingArea.Bottom / 2 - Me.Height / 2

        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New System.Drawing.Point(X, Y)
    End Sub
End Class