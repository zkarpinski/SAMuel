Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Word = Microsoft.Office.Interop.Word
Imports System.IO
Imports System.Drawing.Imaging
Imports System.Drawing.Printing
Imports System.ComponentModel

Public Class frmMain

    Private Sub frmMain_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Application.Exit()
    End Sub
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
        'Welcome the user.
        lblStatus.Text = String.Format("Welcome {0}!", Environment.UserName)

        'Hide DPA Tab for now.
        tabDPA.Dispose()
    End Sub

    Private Sub _DragEnter(ByVal sender As Object, ByVal e As DragEventArgs) Handles tabWordToTiff.DragEnter, tabAddContact.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

#Region "Outlook Tab Region --------------------------------------------------------------------------------------"
    'Outlook Tab wide varibles
    Dim bNextPressed As Boolean
    Dim bRejectPressed As Boolean
    Dim bCancelPressed As Boolean
    Dim bAuditMode As Boolean
    Dim bUnAttendedMode As Boolean

    Dim currentEmailAttachment As String

    Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click
        Dim oApp As Outlook.Application = New Outlook.Application
        Dim wApp As Word.Application = Nothing
        Dim sDestination As String
        Dim sFileExt As String
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

        'Make sure everything is at initial state.
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
        Dim endTime As DateTime
        Dim totalTime As TimeSpan

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
                Else
                    lblOutlookMessage.Text = ""
                End If

                lblStatus.Text = "Downloading attachments..."
                Me.Refresh()
                sEmail.DownloadAttachments()
                'Process each attachment within the email
                If sEmail.AttachmentCount > 0 Then
                    frmEmails.clbSelectedEmails.Items.Add("[" & sEmail.AttachmentCount.ToString & "] " & sEmail.Account & vbTab & vbTab & sEmail.Subject & vbTab & sEmail.From)

                    'Display email info when in auditmode
                    If (bAuditMode) Then
                        Me.Cursor = Cursors.Default
                        lblStatus.Text = "Waiting for user..."
                        Outlook_Setup_Audit_View()
                        txtAcc.Text = sEmail.Account
                        txtSubject.Text = sEmail.Subject
                        txtFrom.Text = sEmail.From
                        rtbEmailBody.Text = sEmail.Body

                        'Fill the listview with all the attachments within the email.
                        lstEmailAttachments.Visible = True
                        For j As Integer = 0 To (sEmail.AttachmentCount - 1)
                            Dim lvi As ListViewItem = New ListViewItem(sEmail.Attachments(j).Filetype)
                            lvi.SubItems.Add(sEmail.Attachments(j).Filename)
                            lvi.Tag = sEmail.Attachments(j).File
                            lstEmailAttachments.Items.Add(lvi)
                        Next
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
                        endTime = DateTime.Now
                        Reset_Outlook_Tab()
                        Reset_ProgressBar()
                        totalTime = endTime - startTime
                        LogAction(0, String.Format("{0}: Finished {1} emails in {2} seconds.", workingOutlookFolder.Parent, completedEmailsCount, totalTime.TotalSeconds))
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    ElseIf bRejectPressed Then
                        '------ LOG reject action? ---------
                    ElseIf (bNextPressed Or bAuditMode = False) And bRejectPressed = False Then
                        'Save the edited attachment as tiff and add to list
                        lblStatus.Text = "Converting  Tiff..."
                        Me.Refresh()
                        lstEmailAttachments.Items.Clear()
                        For j As Integer = 0 To (sEmail.AttachmentCount - 1)

                            Dim currentAttachmentFile As String = sEmail.Attachments(j).File
                            'Verify a valid attachment file type.
                            sFileExt = EmailProcessing.ValidateAttachmentType(currentAttachmentFile)
                            If (sFileExt Is vbNullString) Then
                                'Delete the file and move to next attachment
                                System.IO.File.Delete(sEmail.Attachments(j).File)
                                Continue For
                            End If

                            currentEmailAttachment = sEmail.Attachments(j).File

                            'Name the tiff files to be created.
                            outTiff = [String].Format("{0}{1}_{2}.tiff", sDestination, sEmail.Account, rand.Next(10000).ToString)

                            'Handle how each file type is printed
                            If sFileExt = ".doc" Or sFileExt = ".docx" Then
                                'Convert Word documents to tiff
                                If (wApp Is Nothing) Then
                                    wApp = New Word.Application
                                    wApp = CreateObject("Word.Application")
                                    wApp.WindowState = Word.WdWindowState.wdWindowStateMinimize
                                End If
                                Conversion.docToTiff(currentAttachmentFile, outTiff, wApp)
                            ElseIf sFileExt = ".pdf" Then
                                'Convert PDFs to tiff
                                Conversion.pdfToTiff(currentAttachmentFile, outTiff)
                            Else
                                'Print images to tiff
                                Conversion.imgToTiff(currentAttachmentFile, outTiff)
                            End If

                            'Delete the saved email attachment
                            System.IO.File.Delete(currentAttachmentFile)
                        Next
                    Else
                        'Undesired state. Log this
                        LogAction(99, "Outlook buttonState: " & bCancelPressed.ToString & "," & bRejectPressed.ToString & "," & bNextPressed.ToString & "," & bAuditMode.ToString)
                    End If

                    'Flag email in outlook as complete.
                    sEmail.EmailObject.FlagStatus = Microsoft.Office.Interop.Outlook.OlFlagStatus.olFlagComplete
                    sEmail.EmailObject.Save()

                    sEmail.Dispose()

                    'Checks the email in the check list box
                    frmEmails.clbSelectedEmails.SetItemChecked(ProgressBar.Value, True)
                    completedEmailsCount += 1
                Else
                    'No attachments
                    'TODO Add print option to no attachments case
                    frmEmails.clbSelectedEmails.Items.Add("[0] " & sEmail.Account & vbTab & vbTab & sEmail.Subject & vbTab & sEmail.From)
                End If

            Catch ex As Exception
                'Unknown error with email.
                LogAction(98, String.Format("{0} - {1} : {2}", sEmail.Subject, sEmail.From, ex.ToString))
            End Try

            'Reset variables
            bNextPressed = False
            bRejectPressed = False
            lstEmailAttachments.Items.Clear()
            lblOutlookMessage.Text = ""

            'Update Progress
            ProgressBar.Value += 1
            Me.Refresh()
        Next

        endTime = DateTime.Now
        totalTime = endTime - startTime

        LogAction(0, String.Format("{0}: Finished {1} emails in {2} seconds.", workingOutlookFolder.FolderPath, completedEmailsCount, totalTime.TotalSeconds))

        If completedEmailsCount > 0 Then
            lblStatus.Text = "Finished processing emails."
        Else
            lblStatus.Text = "No emails were processed. See Log."
        End If

        Me.Cursor = Cursors.Default
        Reset_Outlook_Tab()
    End Sub

    Private Sub btnReject_Click(sender As Object, e As EventArgs) Handles btnReject.Click
        bRejectPressed = True
    End Sub

    Private Sub Outlook_Setup_Audit_View()
        groupOLAudit.Visible = True
        rtbEmailBody.Visible = True
        btnReject.Visible = True
        btnCancel.Visible = True
        lstEmailAttachments.Visible = True
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        bCancelPressed = True
    End Sub

    Private Sub btnNext_Click_1(sender As Object, e As EventArgs) Handles btnNext.Click
        ' TODO add validation of account number.
        bNextPressed = True
    End Sub

    Private Sub Reset_Outlook_Tab()
        'Reset the Outlook tab to default
        'variables
        bNextPressed = False
        bRejectPressed = False
        bCancelPressed = False

        'settings
        lstEmailAttachments.Items.Clear()
        btnCancel.Enabled = False
        btnReject.Enabled = False
        btnNext.Visible = False
        btnNext.Enabled = False
        btnRun.Enabled = True
        btnRun.Visible = True
        'Hide audit options
        rtbEmailBody.Visible = False
        btnReject.Visible = False
        btnCancel.Visible = False
        groupOLAudit.Visible = False
        lstEmailAttachments.Visible = False

        'clear text fields
        rtbEmailBody.Text = ""
        txtAcc.Text = ""
        txtFrom.Text = ""
        txtAcc.Text = ""
        txtSubject.Text = ""
        lblOutlookMessage.Text = ""

        Me.Cursor = Cursors.Default
    End Sub

    Private Sub itmEmailAttachment_Click(sender As Object, e As EventArgs) Handles lstEmailAttachments.ItemActivate
        System.Diagnostics.Process.Start(lstEmailAttachments.SelectedItems(0).Tag)
    End Sub

#End Region

#Region "Kofax It Tab Region -------------------------------------------------------------------------------------"
    Private Sub btnKFImport_Click(sender As Object, e As EventArgs) Handles btnKFImport.Click
        Dim batchName As String
        Dim batchType As String
        Dim batchSource As String
        Dim comments As String

        'Setup the progress bar.
        Reset_ProgressBar()
        Me.ProgressBar.Maximum = 1

        'Set values from the form.
        batchName = Me.txtKFBatchName.Text
        batchType = cbKFBatchType.SelectedItem.ToString
        comments = Me.txtKFComments.Text

        'Determine the source type.
        If Me.rbKFEmail.Checked = True Then
            batchSource = "02 - Email"
        Else
            batchSource = "01 - US Mail"
        End If

        'Open the file dialog with tiffs only selectable.
        dlgOpen.Filter = "TIFF Images|*.tif;*.tiff"
        If dlgOpen.ShowDialog() = DialogResult.OK Then
            KofaxModule.CreateXML((dlgOpen.FileNames).ToList, batchName, batchType, batchSource, comments)
        End If

        'Update UI
        Me.lblStatus.Text = "Conversion Complete."
        Me.ProgressBar.Value = 1
        Reset_Kofax_Tab()
    End Sub

    Private Sub Reset_Kofax_Tab()
        'Reset the KofaxIt tab to default
        txtKFBatchName.Text = ""
        txtKFComments.Text = ""
        cbKFBatchType.SelectedIndex = 2
    End Sub
#End Region

#Region "RightFax It Tab Region ----------------------------------------------------------------------------------"
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
#End Region

#Region "Contacts Tab Region -------------------------------------------------------------------------------------"
    'Contacts Tab Wide Variables
    Dim _bStopContacts As Boolean = False

    Private Sub btnCAddContact_Click(sender As Object, e As EventArgs) Handles btnCAddContact.Click
        Dim strAccountNumber As String
        Dim strContact As String

        ResetContacts_Tab()
        Me.btnCAddContact.Enabled = False

        'Set variable values from the form.
        strAccountNumber = Me.mtxtCAccount.Text
        strContact = Me.rtbCContact.Text

        'Call the script
        Try
            AddContact.RunScript(strAccountNumber, strContact)
        Catch ex As FileNotFoundException
            MsgBox("addContact.exe not found.", MsgBoxStyle.Exclamation)
            Me.btnCAddContact.Enabled = True
            Exit Sub
        End Try

        Me.btnCAddContact.Enabled = True
    End Sub

    Private Sub _DragDropContact(sender As Object, e As DragEventArgs) Handles tabAddContact.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            ResetContacts_Tab()

            Dim files As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
            Dim strContact As String = Me.rtbCContact.Text

            Me.btnCAddContact.Enabled = False
            btnStopAddContacts.Enabled = True
            For Each sFile As String In files
                Dim strAcc As String = GlobalModule.RegexAcc(Path.GetFileNameWithoutExtension(sFile), "\d{5}-\d{5}")

                'Skip files without and account number.
                If strAcc = "X" Then
                    lbxContactAlerts.Visible = True
                    lbxContactAlerts.Items.Add(String.Format("{0} was skipped.", Path.GetFileName(sFile)))
                    Me.Refresh()
                    Continue For
                End If

                'Check if user interupted process.
                If (_bStopContacts) Then
                    Me.btnCAddContact.Enabled = True
                    Exit Sub
                End If

                'TODO Make this a Background worker.

                'Call the script
                Try
                    Dim bgContactWorker As BackgroundWorker = New BackgroundWorker
                    AddContact.RunScript(strAcc, strContact)
                Catch ex As FileNotFoundException
                    MsgBox("addContact.exe not found.", MsgBoxStyle.Exclamation)
                    Me.btnCAddContact.Enabled = True
                    Exit Sub
                End Try
            Next

        End If
    End Sub

    Private Sub btnStopAddContacts_Click(sender As Object, e As EventArgs) Handles btnStopAddContacts.Click
        _bStopContacts = True
    End Sub

    Private Sub ResetContacts_Tab()
        btnStopAddContacts.Enabled = False
        _bStopContacts = False
        btnCAddContact.Enabled = True
        rtbCContact.Text = ""
        mtxtCAccount.Text = ""
        lbxContactAlerts.Visible = False
        lbxContactAlerts.Items.Clear()
    End Sub
#End Region

#Region "Convert Tab Region --------------------------------------------------------------------------------------"

    Private Sub btnConvert_Click(sender As Object, e As EventArgs) Handles btnConvert.Click
        'Converts Word Documents to .tif or PDF using MODI
        Reset()

        'If files are selected continue code
        If dlgOpen.ShowDialog() = DialogResult.OK Then
            Try
                If (Me.rbConvertDOC.Checked) Then
                    Conversion.docsToTiff(dlgOpen.FileNames)
                Else
                    Conversion.imagesToTiff(dlgOpen.FileNames)
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

    Private Sub _DragDrop(ByVal sender As Object, ByVal e As DragEventArgs) Handles tabWordToTiff.DragDrop
        'http://www.codemag.com/Article/0407031
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then

            Dim files As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
            Try
                If (Me.rbConvertDOC.Checked) Then
                    Conversion.docsToTiff(files)
                Else
                    Conversion.imagesToTiff(files)
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                Return
            End Try
        End If

    End Sub
#End Region

#Region "DPA Tab Region ------------------------------------------------------------------------------------------"
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

    Private Sub btnBudgetBill_Click(sender As Object, e As EventArgs) Handles btnBudgetBill.Click
        Dim accNumber As String
        accNumber = mtxtDPAAcc.Text
        Call OpenCSSAcc(accNumber)
        Threading.Thread.Sleep(400)
        Call EnrollBB()
    End Sub

    Private Sub btnDPAprocess_Click(sender As Object, e As EventArgs) Handles btnDPAprocess.Click
        'Process the DPA for the desired account
        Dim accNumber As String
        accNumber = mtxtDPAAcc.Text
        Call OpenCSSAcc(accNumber)
        Call OpenPA()
    End Sub

    Private Sub Reset_DPA_Tab()
        'Reset the DPA tab to default
        mtxtDPAAcc.Text = ""
        txtDPAdown.Text = ""
        txtDPAmonthly.Text = ""
        chkMinPayment.Checked = False
    End Sub
#End Region

#Region "Form UI Region ------------------------------------------------------------------------------------------"
    Private Sub Reset_All_Forms()
        'Resets the form state to default
        Reset_Outlook_Tab()
        Reset_DPA_Tab()
        Reset_RightFax_Tab()
        Reset_ProgressBar()
        frmEmails.clbSelectedEmails.Items.Clear()
        Reset_Kofax_Tab()
        ResetContacts_Tab()
    End Sub

    Private Sub Reset_ProgressBar()
        'Progress bar reset
        ProgressBar.Value = 0
        lblStatus.Text = ""
    End Sub

    Private Sub TabControl1_Changed(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        'Resets the form state to default
        Reset_All_Forms()
    End Sub
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        End
    End Sub
    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        frmAbout.Show()
    End Sub
    Private Sub OptionsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OptionsToolStripMenuItem.Click
        frmOptions.Show()
    End Sub
#End Region

End Class