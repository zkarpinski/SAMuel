Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Word = Microsoft.Office.Interop.Word
Imports System.IO
Imports System.Drawing.Imaging
Imports System.Drawing.Printing

Public Class frmMain
    'Form wide variables
    Dim bNextPressed As Boolean
    Dim bRejectPressed As Boolean
    Dim bCancelPressed As Boolean
    Dim bAuditMode As Boolean
    Dim bUnAttendedMode As Boolean

    Public sImageToPrint As String


    Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click
        Dim oApp As Outlook.Application = New Outlook.Application
        Dim wApp As Word.Application
        Dim sDestination As String
        Dim sFile As String, sFileExt As String
        Dim outTiff As String
        Dim emailCount As Integer
        Dim samEmails As New List(Of SAM_Email)
        Dim sEmail As SAM_Email
        Dim rand As New Random
        Dim emailItem As Outlook.MailItem
        Dim completedEmailsCount As Integer = 0

        'Define the save location and check if it exists
        sDestination = My.Settings.savePath + "tiffs\"
        GlobalModule.CheckFolder(sDestination)

        Reset_Outlook_Tab()
        frmEmails.clbSelectedEmails.Items.Clear()
        Reset_ProgressBar()
        frmEmails.Show()

        'enable buttons
        btnCancel.Enabled = True
        btnReject.Enabled = True
        btnNext.Visible = True
        btnNext.Enabled = True
        btnRun.Enabled = False
        btnRun.Visible = False

        Dim workingOutlookFolder As Outlook.MAPIFolder
        workingOutlookFolder = oApp.GetNamespace("MAPI").PickFolder()
        If workingOutlookFolder Is Nothing Then
            Reset_Outlook_Tab()
            Exit Sub
        End If

        Dim startTime As DateTime = DateTime.Now

        'Create each SAM Email
        Try
            For Each emailItem In workingOutlookFolder.Items
                Me.Cursor = Cursors.WaitCursor
                Try
                    lblStatus.Text = "Saving attachments..."
                    Me.Refresh()
                    sEmail = New SAM_Email(emailItem)
                    samEmails.Add(sEmail)
                Catch ex As Exception
                    LogAction(action:="An email was skipped because " & ex.Message)
                    'Move to next email
                    Continue For
                End Try
            Next
        Catch ex As Exception
            MsgBox("Something happened... Sorry", MsgBoxStyle.Exclamation)
            LogAction(action:=ex.Message)
            Reset_Outlook_Tab()
            Me.Cursor = Cursors.Default
            Exit Sub
        End Try

        'Get the email count and update the progress bar
        emailCount = samEmails.Count
        ProgressBar.Maximum = emailCount

        'Update audit and unattended values
        bAuditMode = chkAuditMode.Checked
        bUnAttendedMode = chkValidOnly.Checked

        'Process each email
        For Each sEmail In samEmails
            Try
                'Force an audit if the email is found to be not valid and not in unattendedmode.
                If (Not sEmail.IsValid) And (bUnAttendedMode = False) Then
                    bAuditMode = True
                    lblOutlookMessage.Text = "Invalid Email Found!"
                ElseIf (Not sEmail.IsValid) And (bUnAttendedMode = True) Then
                    'Skip the email
                    frmEmails.clbSelectedEmails.Items.Add("[SKIP] " & sEmail.Subject & vbTab & sEmail.From)
                    ProgressBar.Value += 1
                    Continue For
                End If

                sEmail.DownloadAttachments()

                'Process each attachment within the email
                If sEmail.AttachmentsCount > 0 Then
                    frmEmails.clbSelectedEmails.Items.Add("[" & sEmail.AttachmentsCount.ToString & "] " & sEmail.Account & vbTab & vbTab & sEmail.Subject & vbTab & sEmail.From)
                    For Each sFile In sEmail.Attachments
                        'Verify a valid attachment file type.
                        sFileExt = EmailProcessing.ValidateAttachmentType(sFile)
                        If (sFileExt Is vbNullString) Then
                            'Delete the file and move to next attachment
                            System.IO.File.Delete(sFile)
                            Continue For
                        ElseIf sFileExt = ".pdf" Then
                            ' TODO add PDF handling
                            'EmailProcessing.ParsePDFImgs(sFile)
                            LogAction(0, "PDF detected. " & sEmail.Subject & " " & sEmail.From)
                            Continue For
                        End If

                        sImageToPrint = sFile


                        'Display email info when in auditmode
                        If (bAuditMode) Then
                            Me.Cursor = Cursors.Default
                            lblStatus.Text = "Waiting for user..."
                            Outlook_Setup_Audit_View()
                            txtAcc.Text = sEmail.Account
                            txtSubject.Text = sEmail.Subject
                            txtFrom.Text = sEmail.From
                            rtbEmailBody.Text = sEmail.Body

                            If sFileExt = ".doc" Or sFileExt = ".docx" Then
                                ' TODO Preview word doc
                            Else
                                'Load the attachment
                                picImage.ImageLocation = sImageToPrint
                            End If
                        End If

                        'Wait for user validation of attachment
                        Do Until (bNextPressed = True Or bRejectPressed = True Or bCancelPressed = True Or bAuditMode = False)
                            Application.DoEvents()
                        Loop

                        Me.Cursor = Cursors.WaitCursor

                        'Update sEmail account if its in audit mode
                        If (bAuditMode) Then
                            sEmail.Account = txtAcc.Text
                        End If

                        bAuditMode = chkAuditMode.Checked

                        If bCancelPressed Then
                            'When canceled, release image and reset form and end the routine
                            picImage.Image = Nothing
                            Reset_Outlook_Tab()
                            Reset_ProgressBar()
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        ElseIf bRejectPressed Then
                            '------ LOG reject action? ---------
                        ElseIf (bNextPressed Or bAuditMode = False) And bRejectPressed = False Then
                            'Save the edited attachment as tiff and add to list
                            lblStatus.Text = "Converting to Tiff..."
                            Me.Refresh()

                            'Handle how each file type is printed
                            If sFileExt = ".doc" Or sFileExt = ".docx" Then
                                If (wApp Is Nothing) Then
                                    wApp = New Word.Application
                                    wApp = CreateObject("Word.Application")
                                    wApp.WindowState = Word.WdWindowState.wdWindowStateMinimize
                                End If
                                Dim objWdDoc As Word.Document
                                outTiff = [String].Format("{0}{1}_{2}.tiff", sDestination, sEmail.Account, rand.Next(10000).ToString)

                                'Open the document within word and don't prompt  for conversion.
                                objWdDoc = wApp.Documents.Open(FileName:=sFile, ConfirmConversions:=False)
                                wApp.Visible = False
                                'Print to Tiff
                                objWdDoc.PrintOut(PrintToFile:=True, OutputFileName:=outTiff)

                                objWdDoc.Close()

                            Else

                                outTiff = [String].Format("{0}{1}_{2}.tiff", sDestination, sEmail.Account, rand.Next(10000).ToString)


                                'Print the file to file with MODI
                                PrintDocument1.PrinterSettings.PrintToFile = True
                                PrintDocument1.PrinterSettings.PrintFileName = outTiff
                                PrintDialog1.PrintToFile = True
                                PrintDocument1.Print()

                            End If
                            Me.Refresh()
                        Else
                            'Undesired state. Log this
                            LogAction(99, "Outlook buttonState: " & bCancelPressed.ToString & "," & bRejectPressed.ToString & "," & bNextPressed.ToString & "," & bAuditMode.ToString)
                        End If

                        'Reset variables
                        bNextPressed = False
                        bRejectPressed = False

                        'Release images
                        If Not IsNothing(picImage.Image) Then picImage.Image.Dispose()
                        picImage.Image = Nothing
                        picImage.ImageLocation = vbNullString

                        lblStatus.Text = ""


                        Application.DoEvents()
                        Me.Refresh()

                        'Delete the saved email attachment
                        System.IO.File.Delete(sFile)


                    Next
                    sEmail.EmailObject.FlagStatus = Microsoft.Office.Interop.Outlook.OlFlagStatus.olFlagComplete
                    sEmail.EmailObject.Save()

                    sEmail.Dispose()

                    'Checks the email in the check list box
                    frmEmails.clbSelectedEmails.SetItemChecked(ProgressBar.Value, True)
                    completedEmailsCount += 1
                Else
                    'No attachments
                    frmEmails.clbSelectedEmails.Items.Add("[SKIP] " & sEmail.Account & vbTab & vbTab & sEmail.Subject & vbTab & sEmail.From)
                    LogAction(50, String.Format("{0} - {1} was skipped. No attachments", sEmail.Subject, sEmail.From))
                End If



                bNextPressed = False
                ProgressBar.Value += 1
                Me.Refresh()
            Catch ex As Exception
                'Unknown error with email.
                LogAction(98, String.Format("{0} - {1} : {2}", sEmail.Subject, sEmail.From, ex.ToString))
            End Try
        Next

        Dim endTime As DateTime = DateTime.Now
        Dim totalTime As TimeSpan = endTime - startTime

        LogAction(0, String.Format("Finished {0} emails in {1} time span.", completedEmailsCount, totalTime))

        If completedEmailsCount > 0 Then
            lblStatus.Text = "Finished processing emails."
        Else
            lblStatus.Text = "No emails were processed. See Log."
        End If

        Me.Cursor = Cursors.Default
        Reset_Outlook_Tab()

    End Sub

    Private Sub btnConvert_Click(sender As Object, e As EventArgs) Handles btnConvert.Click
        'Converts Word Documents to .tif or PDF using MODI
        Reset()

        Dim oPrinter As Object

        'If files are selected continue code
        If dlgOpen.ShowDialog() = DialogResult.OK Then
            If rbConvertPDF.Checked = True Then
                Try
                    For Each value In dlgOpen.FileNames
                        oPrinter = New cPDF995
                        oPrinter.fileToPrint = value
                        oPrinter.printFile()
                        oPrinter = Nothing
                    Next
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                    btnConvert.Enabled = True
                    Exit Sub
                End Try
            End If

            Try
                If (Me.rbConvertDOC.Checked) Then
                    ConvertToTiff.WordDocs(dlgOpen.FileNames)
                Else
                    ConvertToTiff.ImageFiles(dlgOpen.FileNames)
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                btnConvert.Enabled = True
                Exit Sub
            End Try
        End If

        'Cleanup
        btnConvert.Enabled = True
    End Sub

    Private Sub btnReject_Click(sender As Object, e As EventArgs) Handles btnReject.Click
        bRejectPressed = True
    End Sub

    Private Sub chkMinPayment_CheckedChanged(sender As Object, e As EventArgs) Handles chkMinPayment.CheckedChanged
        'Enters the minimum payment info when checked.
        If chkMinPayment.Checked Then
            txtDPAdown.Text = "$0.00"
            txtDPAmonthly.Text = "$10.00"
            txtDPAmonthly.Enabled = False
        Else
            txtDPAdown.Text = ""
            txtDPAmonthly.Text = ""
            txtDPAmonthly.Enabled = True
        End If
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblStatus.BackColor = Color.Transparent

        'Resets the form state to default
        Reset_All_Forms()

        'Check initial user settings for a save path
        If My.Settings.savePath = "" Then
            Dim defaultPath As String = Environment.GetEnvironmentVariable("userprofile") & "\SAMuel\"
            My.Settings.savePath = defaultPath
            My.Settings.Save()
        End If

        'Verify all output folders exist
        GlobalModule.InitOutputFolders()

    End Sub

    Private Sub TabControl1_Changed(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        'Resets the form state to default
        Reset_All_Forms()
    End Sub

    Private Sub Reset_All_Forms()
        'Resets the form state to default
        Reset_Outlook_Tab()
        Reset_DPA_Tab()
        Reset_RightFax_Tab()
        Reset_ProgressBar()
        frmEmails.clbSelectedEmails.Items.Clear()
        cbKFBatchType.SelectedIndex = 2
    End Sub

    Private Sub Reset_ProgressBar()
        'Progress bar reset
        ProgressBar.Value = 0
        lblStatus.Text = ""
    End Sub

    Private Sub Reset_RightFax_Tab()
        'Reset the RightFax tab  to default
        txtRFuser.Text = My.Settings.rfUser
        txtRFsvr.Text = My.Settings.rfServer
        txtRFpw.Text = My.Settings.rfPW
        chkRFNTauth.Checked = My.Settings.rfUseNT
        txtRFRecFax.Text = My.Settings.rfRecFax
        txtRFRecName.Text = My.Settings.rfRecName
        chkRFSaveRec.Checked = False
        chkRFCoverSheet.Checked = True
    End Sub

    Private Sub Reset_DPA_Tab()
        'Reset the DPA tab to default
        mtxtDPAAcc.Text = ""
        txtDPAdown.Text = ""
        txtDPAmonthly.Text = ""
        chkMinPayment.Checked = False
    End Sub

    Private Sub Reset_Outlook_Tab()
        'Reset the Outlook tab to default
        'variables
        bNextPressed = False
        bRejectPressed = False
        bCancelPressed = False

        'settings
        picImage.Image = Nothing
        btnCancel.Enabled = False
        btnReject.Enabled = False
        btnNext.Visible = False
        btnNext.Enabled = False
        btnRun.Enabled = True
        btnRun.Visible = True
        'Hide audit options
        rtbEmailBody.Visible = False
        btnReject.Visible = False
        lblSelectedEmails.Visible = False
        btnCancel.Visible = False
        groupOLAudit.Visible = False
        picImage.Enabled = False
        picImage.Visible = False


        'clear text fields
        rtbEmailBody.Text = ""
        txtAcc.Text = ""
        txtFrom.Text = ""
        txtAcc.Text = ""
        txtSubject.Text = ""
        lblOutlookMessage.Text = ""
    End Sub

    Private Sub Outlook_Setup_Audit_View()
        groupOLAudit.Visible = True
        rtbEmailBody.Visible = True
        btnReject.Visible = True
        lblSelectedEmails.Visible = False
        btnCancel.Visible = True
        picImage.Enabled = True
        picImage.Visible = True
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        bCancelPressed = True
    End Sub

    Private Sub btnNext_Click_1(sender As Object, e As EventArgs) Handles btnNext.Click
        ' TODO add validation of account number.
        bNextPressed = True
    End Sub

    Private Sub btnDPAprocess_Click(sender As Object, e As EventArgs) Handles btnDPAprocess.Click
        'Process the DPA for the desired account
        Dim accNumber As String
        accNumber = mtxtDPAAcc.Text
        Call OpenCSSAcc(accNumber)
        Call OpenPA()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        End
    End Sub

    Private Sub btnBudgetBill_Click(sender As Object, e As EventArgs) Handles btnBudgetBill.Click
        Dim accNumber As String
        accNumber = mtxtDPAAcc.Text
        Call OpenCSSAcc(accNumber)
        Threading.Thread.Sleep(400)
        Call EnrollBB()
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        frmAbout.Show()
    End Sub

    Private Sub btnRFax_Click(sender As Object, e As EventArgs) Handles btnRFax.Click
        Dim strServerName As String, strUsername As String, strPassword As String 'RightFax Server Strings
        Dim strRecName As String, strRecFax As String 'Fax Recipient Strings
        Dim bUseNTAuth As Boolean, bSaveRec As Boolean, bUseCoverSheet As Boolean
        Dim sFile As String
        Dim i As Integer = 0, faxesCount As Integer

        Dim objRightFax As RFCOMAPILib.FaxServer
        Dim objFax As RFCOMAPILib.Fax

        Reset_ProgressBar()

        'Set values
        strServerName = txtRFsvr.Text
        strUsername = txtRFuser.Text
        strPassword = txtRFpw.Text
        If chkRFNTauth.Checked = True Then
            bUseNTAuth = True
        Else
            bUseNTAuth = False
        End If
        strRecName = txtRFRecName.Text
        strRecFax = txtRFRecFax.Text
        If chkRFSaveRec.Checked = True Then
            bSaveRec = True
        Else
            bSaveRec = False
        End If
        If chkRFCoverSheet.Checked = True Then
            bUseCoverSheet = True
        Else
            bUseCoverSheet = False
        End If

        'Update RF Server User App Settings
        My.Settings.rfUseNT = bUseNTAuth
        My.Settings.rfServer = strServerName
        My.Settings.rfUser = strUsername
        My.Settings.rfPW = strPassword
        'Update RF Recipient User App Settings
        If bSaveRec Then
            My.Settings.rfRecName = strRecName
            My.Settings.rfRecFax = strRecFax
        End If
        My.Settings.Save()


        'Fax the selected files
        If dlgOpen.ShowDialog() = DialogResult.OK Then
            objRightFax = RightFax.ConnectToServer(strServerName, strUsername, bUseNTAuth)
            objRightFax.OpenServer()
            faxesCount = dlgOpen.FileNames.Count
            ProgressBar.Maximum = faxesCount
            'Create and send each fax
            For Each sFile In dlgOpen.FileNames
                i += 1
                lblStatus.Text = String.Format("Faxing {0} of {1}.", i, faxesCount)
                Me.Refresh()
                objFax = RightFax.CreateFax(objRightFax, strRecName, strRecFax, sFile)
                RightFax.SendFax(objFax)
                RightFax.MoveFaxedFile(sFile)
                ProgressBar.Value += 1
                objFax = Nothing
            Next
            lblStatus.Text = "DONE!"
            objRightFax.CloseServer()
        End If
    End Sub

    Private Sub OptionsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OptionsToolStripMenuItem.Click
        frmOptions.Show()
    End Sub

    Private Sub _DragDrop(ByVal sender As Object, ByVal e As DragEventArgs) Handles tabWordToTiff.DragDrop
        'http://www.codemag.com/Article/0407031
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then

            Dim files As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
            Try
                If (Me.rbConvertDOC.Checked) Then
                    ConvertToTiff.WordDocs(files)
                Else
                    ConvertToTiff.ImageFiles(files)
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                Return
            End Try
        End If

    End Sub

    Private Sub _DragEnter(ByVal sender As Object, ByVal e As DragEventArgs) Handles tabWordToTiff.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub btnKFImport_Click(sender As Object, e As EventArgs) Handles btnKFImport.Click
        Dim batchName As String
        Dim batchType As String
        Dim batchSource As String
        Dim comments As String

        batchName = Me.txtKFBatchName.Text
        batchType = cbKFBatchType.SelectedItem.ToString
        comments = Me.txtKFComments.Text

        If Me.rbKFEmail.Checked = True Then
            batchSource = "02 - Email"
        Else
            batchSource = "01 - US Mail"
        End If

        dlgOpen.Filter = "TIFF Images|*.tif;*.tiff"
        If dlgOpen.ShowDialog() = DialogResult.OK Then
            KofaxModule.CreateXML((dlgOpen.FileNames).ToList, batchName, batchType, batchSource, comments)
        End If

    End Sub

    Private Sub btnCAddContact_Click(sender As Object, e As EventArgs) Handles btnCAddContact.Click
        Dim strAccountNumber As String
        Dim strContact As String

        strAccountNumber = Me.mtxtCAccount.Text
        strContact = rtbCContact.Text

        AddContacts.MiscCollection(strAccountNumber, strContact)
    End Sub

    Private Sub tabWordToTiff_Click(sender As Object, e As EventArgs) Handles tabWordToTiff.Click

    End Sub

    Private Sub PrintDocument1_PrintPage(sender As Object, e As PrintPageEventArgs) Handles PrintDocument1.PrintPage
        'Resize rotate image if needed then print within page bounds.
        Using mBitmap As Bitmap = Bitmap.FromFile(sImageToPrint)
            Using bm As Bitmap = ImageProcessing.ResizeImage(mBitmap)
                e.Graphics.DrawImage(CType(bm, Image), 0, 0, e.PageBounds.Width, e.PageBounds.Height)
            End Using
        End Using
    End Sub

    Private Sub picImage_Click(sender As Object, e As EventArgs) Handles picImage.Click
        System.Diagnostics.Process.Start(Me.picImage.ImageLocation.ToString)
    End Sub
End Class