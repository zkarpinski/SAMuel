Imports System.IO
Imports SAMuel.Modules

Public Class FrmOptions

    Private Sub btnSelectSaveFolder_Click(sender As Object, e As EventArgs) Handles btnSelectSaveFolder.Click
        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            Me.txtSavePath.Text = FolderBrowserDialog1.SelectedPath + "\"
        End If
    End Sub

    Private Sub btnSaveOptions_Click(sender As Object, e As EventArgs) Handles btnSaveOptions.Click
        'Update settings, save and close
        If Me.txtSavePath.Text <> My.Settings.savePath Then
            My.Settings.savePath = Me.txtSavePath.Text
            My.Settings.Save()
            GlobalModule.LogAction(actionCode:=2, action:=Me.txtSavePath.Text)
            GlobalModule.InitOutputFolders(My.Settings.savePath)
        End If

        If Not (Me.chkAuditMode.Checked = My.Settings.Audit_Each_Email) Then
            My.Settings.Audit_Each_Email = Me.chkAuditMode.Checked
            My.Settings.Save()
        End If

        Me.Close()
    End Sub

    Private Sub frmOptions_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Load settings into corresponding fields.
        Me.txtSavePath.Text = My.Settings.savePath
        Me.chkAuditMode.Checked = My.Settings.Audit_Each_Email
    End Sub

    Private Sub btnEmpty_Click(sender As Object, e As EventArgs) Handles btnEmpty.Click

        'Delete the save folders and it's contents.
        If (Directory.Exists(ATT_FOLDER)) Then Directory.Delete(ATT_FOLDER, True)
        If (Directory.Exists(FAXED_FOLDER)) Then Directory.Delete(FAXED_FOLDER, True)
        If (Directory.Exists(CONV_FOLDER)) Then Directory.Delete(CONV_FOLDER, True)
        If (Directory.Exists(EMAILS_FOLDER)) Then Directory.Delete(EMAILS_FOLDER, True)

        'Recreate all the folders.
        GlobalModule.InitOutputFolders(My.Settings.savePath)
        MsgBox("SAMuel save folder cleared!", MsgBoxStyle.OkOnly)
    End Sub

    Private Sub btnViewLog_Click(sender As Object, e As EventArgs) Handles btnViewLog.Click
        If (File.Exists(My.Settings.savePath + "SAMuel.log")) Then
            Process.Start(My.Settings.savePath + "SAMuel.log")


        Else
            MsgBox("Logfile not found!", MsgBoxStyle.Exclamation)
        End If

    End Sub

    Private Sub btnShow_Click(sender As Object, e As EventArgs) Handles btnShow.Click
        'Open the saved path folder with windows explorer.
        Process.Start("explorer.exe", My.Settings.savePath)
    End Sub
End Class