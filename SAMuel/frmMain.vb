Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Word = Microsoft.Office.Interop.Word
Imports System.IO

Public Class frmMain

    Dim btnNextPressed As Boolean
    Dim btnRejectPressed As Boolean
    Dim btnRunState As Boolean

    Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click
        Dim oApp As Outlook.Application = New Outlook.Application
        Dim objFolder As Outlook.MAPIFolder = oApp.Session.Folders.GetFirst
        objFolder = objFolder.Folders.Item("Inbox")
        Dim oItems As Outlook.Items = objFolder.Items
        Dim oMsg As Outlook.MailItem
        Dim oAtt As Outlook.Attachment
        Dim sFile As String
        Dim sFileExt As String

        For Each oMsg In oItems
            Debug.Print(oMsg.Subject)
            txtSubject.Text = oMsg.Subject
            txtFrom.Text = oMsg.SenderName
            txtAcc.Text = regexAccount(oMsg.Subject)
            rtbEmailBody.Text = oMsg.Body
            Debug.Print(oMsg.Attachments.Count)
            If oMsg.Attachments.Count > 0 Then
                For Each oAtt In oMsg.Attachments
                    sFile = Environment.GetEnvironmentVariable("userprofile") & "\Desktop\" & oAtt.FileName
                    sFileExt = Path.GetExtension(sFile).ToLower

                    If sFileExt = ".tiff" Or sFileExt = ".png" Or _
                            sFileExt = ".jpg" Or sFileExt = ".jpeg" Or _
                            sFileExt = ".tif" Or sFileExt = ".gif" Then
                        oAtt.SaveAsFile(sFile)
                        picImage.Image = New Bitmap(sFile)
                        Do Until (btnNextPressed = True)
                            Application.DoEvents()
                        Loop
                        btnNextPressed = False
                    End If
                Next
            End If
            Do Until (btnNextPressed = True)
                Application.DoEvents()
            Loop
            btnNextPressed = False
        Next
        'If dlgOpen.ShowDialog() = DialogResult.OK Then
        'picImage.Image = New Bitmap(dlgOpen.FileName)
        'End If
    End Sub

    Private Sub btnConvert_Click(sender As Object, e As EventArgs) Handles btnConvert.Click
        Dim objWdDoc As Word.Document
        Dim objWord As Word.Application
        Dim sDoc As String
        Dim sDestination As String = Environment.GetEnvironmentVariable("userprofile") & "\Desktop\"
        Dim sFileName As String

        Form_Reset()


        'If files are selected continue code
        If dlgOpen.ShowDialog() = DialogResult.OK Then
            Dim i As Integer
            i = 0
            objWord = CreateObject("Word.Application")
            objWord.ActivePrinter = "Fax"
            ProgressBar.Maximum = dlgOpen.FileNames.Count()
            For Each sDoc In dlgOpen.FileNames

                sFileName = Path.GetFileNameWithoutExtension(sDoc)
                i += 1
                objWdDoc = objWord.Documents.Open(sDoc)
                objWord.Visible = False
                'Print to Tiff
                objWdDoc.PrintOut(PrintToFile:=True, _
                                  OutputFileName:=sDestination & sFileName & ".tiff")
                objWdDoc.Close()
                ProgressBar.Value += 1
            Next
            lblDone.Visible = True
            objWord.Quit()
        End If

        'General Cleanup
        btnConvert.Enabled = True
        objWdDoc = Nothing
        objWord = Nothing
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        btnNextPressed = True
    End Sub

    Private Sub chkMinPayment_CheckedChanged(sender As Object, e As EventArgs) Handles chkMinPayment.CheckedChanged
        If chkMinPayment.Checked Then
            txtDPAdown.Text = "$0.00"
            txtDPAmonthly.Text = "$10.00"
            txtDPAdown.Enabled = False
            txtDPAmonthly.Enabled = False
        Else
            txtDPAdown.Text = ""
            txtDPAmonthly.Text = ""
            txtDPAdown.Enabled = True
            txtDPAmonthly.Enabled = True
        End If
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Form_Reset()
    End Sub

    Private Sub TabControl1_Changed(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        form_Reset()

    End Sub

    Private Sub Form_Reset()
        'Outlook Tab
        btnNextPressed = False
        btnRejectPressed = False
        btnRunState = False


        'Progress bar reset
        ProgressBar.Value = 0
        lblDone.Visible = False
    End Sub
End Class
