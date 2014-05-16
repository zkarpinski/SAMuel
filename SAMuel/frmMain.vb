Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Word = Microsoft.Office.Interop.Word
Imports System.IO

Public Class frmMain

    Dim btnNextPressed As Boolean
    Dim btnRejectPressed As Boolean
    Dim btnCancelPressed As Boolean

    Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click
        Dim oApp As Outlook.Application = New Outlook.Application
        Dim objFolder As Outlook.MAPIFolder = oApp.Session.Folders.GetFirst
        objFolder = objFolder.Folders.Item("Inbox")
        Dim oItems As Outlook.Items = objFolder.Items
        Dim oMsg As Outlook.MailItem
        Dim oAtt As Outlook.Attachment
        Dim sDestination As String = Environment.GetEnvironmentVariable("userprofile") & "\Desktop\"
        Dim sFile As String
        Dim sFileExt As String

        'enable buttons
        btnCancel.Enabled = True
        btnReject.Enabled = True
        btnNext.Visible = True
        btnNext.Enabled = True
        btnRun.Enabled = False
        btnRun.Visible = False

        'For each email in source folder
        For Each oMsg In oItems
            'Fill out textbox values
            Debug.Print(oMsg.Subject)
            txtSubject.Text = oMsg.Subject
            txtFrom.Text = oMsg.SenderName
            txtAcc.Text = regexAccount(oMsg.Subject)
            rtbEmailBody.Text = oMsg.Body
            Debug.Print(oMsg.Attachments.Count)

            If oMsg.Attachments.Count > 0 Then
                'For every attachment within the email
                For Each oAtt In oMsg.Attachments
                    sFile = sDestination & oAtt.FileName
                    'Verify a valid attachment file type
                    sFileExt = Path.GetExtension(sFile).ToLower
                    If sFileExt = ".tiff" Or sFileExt = ".png" Or _
                            sFileExt = ".jpg" Or sFileExt = ".jpeg" Or _
                            sFileExt = ".tif" Or sFileExt = ".gif" Then
                        'Save the attachment then load the preview
                        oAtt.SaveAsFile(sFile)
                        picImage.Image = New Bitmap(sFile)
                        'Wait for user validation of attachment
                        Do Until (btnNextPressed = True Or btnRejectPressed = True Or btnCancelPressed = True)
                            Application.DoEvents()
                        Loop
                        If btnCancelPressed Then
                            'When canceled, reset form and end the routine
                            Form_Reset()
                            Exit Sub
                        ElseIf btnNextPressed Then
                            '------- ADD CONVERT TO TIFF ------------

                            '------- ADD WATERMARK OF ACCOUNT # TO TIFF ---------
                        ElseIf btnRejectPressed Then
                            '------ LOG reject action? ---------
                        End If

                        'Reset variables
                        btnNextPressed = False
                        btnRejectPressed = False

                        'Release image
                        picImage.Image.Dispose()
                        picImage.Image = Nothing
                        'Delete the saved email attachment
                        ''System.IO.Directory.GetFiles() ' Get all files within a folder ** useful later **
                        System.IO.File.Delete(sFile)
                    End If
                Next
            End If
            'Wait for user to move to next email
            Do Until (btnNextPressed = True)
                Application.DoEvents()
            Loop
            btnNextPressed = False
        Next
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
            'Initiate word application object
            objWord = CreateObject("Word.Application")
            'Set active printer to Fax
            objWord.ActivePrinter = "Fax"
            'Determine progressbar maximum value from number of files chosen
            ProgressBar.Maximum = dlgOpen.FileNames.Count()
            For Each sDoc In dlgOpen.FileNames
                'Get the file name
                sFileName = Path.GetFileNameWithoutExtension(sDoc)
                'Open the document within word
                objWdDoc = objWord.Documents.Open(sDoc)
                objWord.Visible = False
                'Print to Tiff
                objWdDoc.PrintOut(PrintToFile:=True, _
                                  OutputFileName:=sDestination & sFileName & ".tiff")
                'Release document
                objWdDoc.Close()
                'Progress the progress bar
                ProgressBar.Value += 1
            Next
            lblDone.Visible = True
            objWord.Quit()
        End If

        'Cleanup
        btnConvert.Enabled = True
        objWdDoc = Nothing
        objWord = Nothing
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnReject.Click
        btnRejectPressed = True
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
        'Resets the form state to default
        Form_Reset()
    End Sub

    Private Sub TabControl1_Changed(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        'Resets the form state to default
        Form_Reset()
    End Sub

    Private Sub Form_Reset()
        'Resets the form state to default

        'Outlook Tab
        'variables
        btnNextPressed = False
        btnRejectPressed = False
        btnCancelPressed = False
        'settings
        picImage.Image = Nothing
        btnCancel.Enabled = False
        btnReject.Enabled = False
        btnNext.Visible = False
        btnNext.Enabled = False
        btnRun.Enabled = True
        btnRun.Visible = True



        'Progress bar reset
        ProgressBar.Value = 0
        lblDone.Visible = False
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        btnCancelPressed = True
    End Sub

    Private Sub btnNext_Click_1(sender As Object, e As EventArgs) Handles btnNext.Click
        btnNextPressed = True
    End Sub
End Class
