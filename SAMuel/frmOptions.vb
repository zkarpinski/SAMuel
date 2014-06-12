Public Class frmOptions

    Private Sub btnSelectSaveFolder_Click(sender As Object, e As EventArgs) Handles btnSelectSaveFolder.Click
        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            Me.txtSavePath.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub btnSaveOptions_Click(sender As Object, e As EventArgs) Handles btnSaveOptions.Click
        'Update settings, save and close
        My.Settings.savePath = Me.txtSavePath.Text
        My.Settings.Save()
        Me.Close()
    End Sub

    Private Sub frmOptions_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.txtSavePath.Text = My.Settings.savePath
    End Sub
End Class