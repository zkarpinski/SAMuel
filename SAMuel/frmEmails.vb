Public Class FrmEmails

    Private Sub frmEmails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Center Form
        Dim mainScreen As Screen = Screen.FromPoint(Me.Location)
        Dim x As Integer = (mainScreen.WorkingArea.Width - Me.Width)
        Dim y As Integer = mainScreen.WorkingArea.Bottom / 2 - Me.Height / 2

        Me.StartPosition = FormStartPosition.Manual
        Me.Location = New Point(x, y)
    End Sub
End Class