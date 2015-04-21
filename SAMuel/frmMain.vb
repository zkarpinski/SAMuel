Imports System.IO
Imports RFCOMAPILib
Imports SAMuel.Classes
Imports Microsoft.Office.Interop.Outlook
Imports SAMuel.Modules
Imports Microsoft.Office.Interop.Word
Imports System.Threading

Public Class FrmMain
    Private Sub frmMain_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        System.Windows.Forms.Application.Exit()
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Resets the form state to default
        Reset_All_Forms()

        'Check initial user settings for a save path
        If My.Settings.savePath = "" Or My.Settings.savePath = "DEFAULT" Then
			'Dim defaultPath As String = Environment.GetEnvironmentVariable("userprofile") & "\SAMuel\"
			Dim defaultPath As String = "C:\SAMuel\"
			My.Settings.savePath = defaultPath
			My.Settings.Save()
		End If

        'Verify all output folders exist
        InitOutputFolders(My.Settings.savePath)
        'Welcome the user.
        lblStatus.Text = String.Format("Welcome {0}!", Environment.UserName)

        'Hide DPA Tab for now.
        tabDPA.Dispose()
    End Sub


#Region "Outlook Tab Region --------------------------------------------------------------------------------------"
    'Outlook Tab wide variables
    Dim _bNextPressed As Boolean
    Dim _bRejectPressed As Boolean
    Dim _bCancelPressed As Boolean
    Dim _bAuditMode As Boolean
    Dim _bThrowAudit As Boolean
    Dim _bUnAttendedMode As Boolean

    Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click
        Dim oApp As Microsoft.Office.Interop.Outlook.Application = New Microsoft.Office.Interop.Outlook.Application
        Dim wApp As Microsoft.Office.Interop.Word.Application = Nothing
        Dim sTiffDestination As String
        Dim sFileExt As String
        Dim outTiff As String
        Dim samEmails As List(Of SamEmail)

        Dim sEmail As SamEmail
        Dim rand As New Random
		Dim completedEmailsCount As Integer = 0
		Dim consecutiveErrors As Short

        'Make sure everything is at initial state.
        Reset_Outlook_Tab()

		Reset_ProgressBar()


        'enable buttons
        btnCancel.Enabled = True
        btnReject.Enabled = True
        btnNext.Visible = True
        btnNext.Enabled = True
        btnRun.Enabled = False
		btnRun.Visible = False
		Me.Refresh()

        'User chooses the folder to work from.
        Dim workingOutlookFolder As MAPIFolder
        workingOutlookFolder = oApp.GetNamespace("MAPI").PickFolder()
        If workingOutlookFolder Is Nothing Then
            Reset_Outlook_Tab()
            Exit Sub
        End If

        Dim startTime As DateTime = DateTime.Now
        Dim endTime As DateTime
        Dim totalTime As TimeSpan

        'Grab the emails and process them into SAM_Emails
        Try
            lblStatus.Text = "Grabbing emails..."
            Me.Cursor = Cursors.WaitCursor
            Me.Refresh()
            samEmails = GetEmails(workingOutlookFolder)
        Catch ex As System.Exception
            MsgBox("Something happened... Sorry", MsgBoxStyle.Exclamation)
            LogAction(action:=ex.Message)
            Reset_Outlook_Tab()
            Exit Sub
        End Try

        'Get the email count and update the progress bar
        Dim emailCount As Integer = samEmails.Count
        ProgressBar.Maximum = emailCount

		'Update audit and unattended values
		'_bAuditMode = My.Settings.Audit_Each_Email ''DISABLED 03/24/15
		_bAuditMode = False
		_bUnAttendedMode = chkValidOnly.Checked
        chkValidOnly.Visible = False

        'Define the save location and check if it exists
        sTiffDestination = EMAILS_FOLDER
        CheckFolder(sTiffDestination)

        'Process each email
        For Each sEmail In samEmails
            Try
                'Force an audit if the email is found to be not valid and not in unattended mode.
                If (Not sEmail.IsValid) And (_bUnAttendedMode = False) Then
					_bThrowAudit = True
					If sEmail.Accounts.Length = 0 Then
						lblOutlookMessage.Text = "No account(s) found!"
					Else
						lblOutlookMessage.Text = "Invalid Email Found!"
					End If

				ElseIf (Not sEmail.IsValid) And (_bUnAttendedMode) Then
                    'Skip the email
                    ProgressBar.Value += 1
                    Continue For
                Else
					lblOutlookMessage.Text = vbNullString
				End If

                lblStatus.Text = "Downloading attachments..."
                Me.Refresh()
                sEmail.DownloadAttachments(ATT_FOLDER)
                'Process each attachment within the email
                If sEmail.AttachmentCount > 0 Then
                    'Display email info when in auditmode or audit thrown.
                    If (_bAuditMode Or _bThrowAudit) Then
						Me.Cursor = Cursors.Default
						lblStatus.Text = "Waiting for user..."
						Outlook_Setup_Audit_View()
						'Handle account view
						If sEmail.Accounts Is Nothing Then
							txtAcc.Text = vbNullString
						ElseIf sEmail.Accounts.Length <= 0 Then
							txtAcc.Text = vbNullString
						Else
							txtAcc.Text = sEmail.Accounts(0)
						End If
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

                        'Wait for user interaction
                        Do Until (_bNextPressed Or _bRejectPressed Or _bCancelPressed)
							System.Windows.Forms.Application.DoEvents()
						Loop
					End If

					Me.Cursor = Cursors.WaitCursor

                    'Update sEmail account if its in audit mode
                    If (_bAuditMode Or _bThrowAudit) Then
						sEmail.Accounts = Split(txtAcc.Text, ";")
					End If

					If _bCancelPressed Then
                        'When canceled, reset form and end the routine
                        endTime = DateTime.Now
						Reset_Outlook_Tab()
						Reset_ProgressBar()
						totalTime = endTime - startTime

                        'Delete the temp attachment folder, then recreate it.
                        DeleteSavedAttachments(ATT_FOLDER)
						CheckFolder(ATT_FOLDER)

						LogAction(0,
								  String.Format("{0}: Finished {1} emails in {2} seconds.", workingOutlookFolder.Parent,
												completedEmailsCount, totalTime.TotalSeconds))
						Exit Sub
					ElseIf _bRejectPressed Then
                        '------ LOG reject action? ---------
                    ElseIf (_bNextPressed Or _bAuditMode = False) And _bRejectPressed = False Then
                        'Save the edited attachment as tiff and add to list
                        lblStatus.Text = "Converting  Tiff..."
						Me.Refresh()
						lstEmailAttachments.Items.Clear()
						For j As Integer = 0 To (sEmail.AttachmentCount - 1)
							Dim currentAttachmentFile As String = sEmail.Attachments(j).File
							'Verify a valid attachment file type.
							sFileExt = ValidateAttachmentType(currentAttachmentFile)
							If (sFileExt Is vbNullString) Then
								'Delete the file and move to next attachment
								File.Delete(sEmail.Attachments(j).File)
								Continue For
							End If

							'Convert the attachment for each account number that was found.
							For k As Integer = 0 To sEmail.Accounts.Length - 1

								'Name the tiff files to be created.
								outTiff = [String].Format("{0}{1}_{2}.tiff", sTiffDestination, sEmail.Accounts(k),
													  rand.Next(10000).ToString)

								'Handle how each file type is printed
								If sFileExt = ".doc" Or sFileExt = ".docx" Then	'Convert Word documents to tiff
									'Open an instance of word if one wasn't created already.
									If (wApp Is Nothing) Then
										wApp = New Microsoft.Office.Interop.Word.Application
										wApp = CreateObject("Word.Application")
										wApp.WindowState = WdWindowState.wdWindowStateMinimize
									End If
									DocToTiff(currentAttachmentFile, outTiff, wApp)
								ElseIf sFileExt = ".pdf" Then 'Convert PDFs to tiff
									PdfToTiff(currentAttachmentFile, outTiff)
								Else 'Print images to tiff
									ImgToTiff(currentAttachmentFile, outTiff)
								End If
							Next
						Next

						'Flag email in outlook as complete.
						sEmail.EmailObject.FlagStatus = OlFlagStatus.olFlagComplete
						sEmail.EmailObject.Save()

						sEmail.Dispose()

						completedEmailsCount += 1
						consecutiveErrors = 0

					Else
                        'Undesired state. Log this
                        LogAction(99,
								  "Outlook buttonState: " & _bCancelPressed.ToString & "," & _bRejectPressed.ToString &
								  "," & _bNextPressed.ToString & "," & _bAuditMode.ToString)
						consecutiveErrors += 1
					End If

				Else
                    'No attachments
                    'TODO Add print option to no attachments case
                End If

            Catch ex As System.Exception
                'Unknown error with email.
                LogAction(98, String.Format("{0} - {1} : {2}", sEmail.Subject, sEmail.From, ex.ToString))
				consecutiveErrors += 1
			End Try

            'Reset variables
            _bNextPressed = False
            _bRejectPressed = False
            _bThrowAudit = False
            lstEmailAttachments.Items.Clear()
			lblOutlookMessage.Text = vbNullString

            'Update Progress
            ProgressBar.Value += 1
			Me.Refresh()

			'Stop if too many errors occured
			If consecutiveErrors >= 5 Then
				MsgBox("Too many consectutive emails failed/had errors. Please check the log and/or try again.", vbCritical)
				LogAction(0, "Too many consectutive emails failed/had errors.")
				Exit For
			End If
		Next

        'Delete the temp attachment folder, then recreate it.
        DeleteSavedAttachments(ATT_FOLDER)
        CheckFolder(ATT_FOLDER)

        'Metrics
        endTime = DateTime.Now
        totalTime = endTime - startTime

        LogAction(0,
                  String.Format("{0}: Finished {1} emails in {2} seconds.", workingOutlookFolder.FolderPath,
                                completedEmailsCount, totalTime.TotalSeconds))

        If completedEmailsCount > 0 Then
			lblStatus.Text = String.Format("Finished processing {0} emails.", completedEmailsCount)
		Else
            lblStatus.Text = "No emails were processed. See Log."
        End If

        Reset_Outlook_Tab()
    End Sub

    Private Sub btnReject_Click(sender As Object, e As EventArgs) Handles btnReject.Click
        _bRejectPressed = True
    End Sub

    Private Sub Outlook_Setup_Audit_View()
        groupOLAudit.Visible = True
        rtbEmailBody.Visible = True
        btnReject.Visible = True
        btnCancel.Visible = True
        lstEmailAttachments.Visible = True
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        _bCancelPressed = True
    End Sub

	Private Sub btnNext_Click_1(sender As Object, e As EventArgs) Handles btnNext.Click
		If ContainsAccountOrCustomer(txtAcc.Text) Then
			' TODO add validation of account number.
			_bNextPressed = True
		Else
			MsgBox("Please enter an account, customer number or press reject to continue.")
		End If
	End Sub

	Private Shared Sub ListViewItemActivate_Open(sender As Object, e As EventArgs) _
        Handles lstEmailAttachments.ItemActivate
        Process.Start(sender.SelectedItems(0).Tag)
    End Sub

    Private Sub Reset_Outlook_Tab()
        'Reset the Outlook tab to default view and values
        'variables
        _bNextPressed = False
        _bRejectPressed = False
        _bCancelPressed = False
        _bThrowAudit = False

        'settings
        lstEmailAttachments.Items.Clear()
        btnCancel.Enabled = False
        btnReject.Enabled = False
        btnNext.Visible = False
        btnNext.Enabled = False
        btnRun.Enabled = True
        btnRun.Visible = True
        chkValidOnly.Visible = True
        'Hide audit options
        rtbEmailBody.Visible = False
        btnReject.Visible = False
        btnCancel.Visible = False
        groupOLAudit.Visible = False
        lstEmailAttachments.Visible = False

        'clear text fields
        rtbEmailBody.Text = vbNullString
		txtAcc.Text = vbNullString
		txtFrom.Text = vbNullString
		txtAcc.Text = vbNullString
		txtSubject.Text = vbNullString
		lblOutlookMessage.Text = vbNullString

		Me.Cursor = Cursors.Default
    End Sub

#End Region

#Region "Kofax Import Tab Region -------------------------------------------------------------------------------------"
	Private kfMode As Byte = 0
	Private kfAccounts() As String
	Private Sub KofaxImportModeChange(sender As Object, e As EventArgs) Handles rbKFStandard.CheckedChanged, rbKFMultiAcc.CheckedChanged, RadioButton1.CheckedChanged
		If TabControl1.SelectedTab.Text = "Kofax Import" Then
			kfMode = (gbKFOptions.Controls.OfType(Of RadioButton).Where(Function(r) r.Checked = True).FirstOrDefault()).Tag
			Select Case (kfMode)
				Case 0 'Standard
					gbKFAccounts.Visible = False
					lblKFCopies.Visible = False
					txtKFNumCopies.Visible = False
				Case 1 'Load Account List
					gbKFAccounts.Visible = True
					lblKFCopies.Visible = False
					txtKFNumCopies.Visible = False
				Case 2 'Multiple Copies
					gbKFAccounts.Visible = False
					lblKFCopies.Visible = True
					txtKFNumCopies.Visible = True
				Case Else
					gbKFAccounts.Visible = False
					lblKFCopies.Visible = False
					txtKFNumCopies.Visible = False
			End Select
		End If
	End Sub

	Private Sub btnKFLoadList_Click(sender As Object, e As EventArgs) Handles btnKFLoadList.Click
		'Clear Listbox
		lbKFAccounts.Items.Clear()
		gbKFAccounts.Text = "Accounts: "
		'Open the file dialog with txt only selectable.
		Dim dlgKFAccounts As OpenFileDialog = New OpenFileDialog
		dlgKFAccounts.Filter = "Text Files (*.txt)|*.txt"
		dlgKFAccounts.Multiselect = False
		If dlgKFAccounts.ShowDialog() = DialogResult.OK Then
			'Read the textfile and get an array of all the account numbers found.
			Using sr As StreamReader = New StreamReader(dlgKFAccounts.FileName)
				kfAccounts = RegexAccCollection(sr.ReadToEnd(), ACCOUNT_REGEX_GROUPED_FORMAT)
			End Using
			gbKFAccounts.Text = "Accounts: " & kfAccounts.Count
			'Add each account to the listbox
			If kfAccounts.Count >= 1 Then
				For Each sAccount As String In kfAccounts
					lbKFAccounts.Items.Add(sAccount)
				Next
			Else
				kfAccounts = New String() {}
				'No Accounts Found
				'Show message
			End If
		End If

	End Sub

	Private Sub btnKFImport_Click(sender As Object, e As EventArgs) Handles btnKFImport.Click
        Dim batchName As String
        Dim batchType As String
        Dim batchSource As String
		Dim comments As String

		Select Case (kfMode)
			Case 1
				If kfAccounts Is Nothing Then
					MsgBox("Load a text file, containing account numbers, before proceeding.", MsgBoxStyle.Exclamation, "Load Account List Mode: Error")
					Exit Sub
				ElseIf kfAccounts.Count <= 0 Then
					MsgBox("Load a text file, containing account numbers, before proceeding.", MsgBoxStyle.Exclamation, "Load Account List Mode: Error")
					Exit Sub
				End If
		End Select

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
			Select Case (kfMode)
				Case 0 ' Standard
					CreateXML((dlgOpen.FileNames).ToList, batchName, batchType, batchSource, comments)
					Me.lblStatus.Text = dlgOpen.FileNames.Count & " files ready to be imported by Kofax."
				Case 1 'Load Account List
					CreateAccListXML((dlgOpen.FileNames).ToList, batchName, batchType, batchSource, kfAccounts, comments)
					Me.lblStatus.Text = dlgOpen.FileNames.Count & " files ready to be imported by Kofax."
				Case 2 ' Multiple Copies
					If dlgOpen.FileNames.Count > 1 Then
						MsgBox("This mode can only be used for one file. Please select a different mode or try again.", MsgBoxStyle.Exclamation, "Multiple Copies Mode: Error")
						Exit Sub
					ElseIf dlgOpen.FileNames.Count = 1 Then
						CreateFileCopiesXML(dlgOpen.FileName, batchName, batchType, batchSource, txtKFNumCopies.Value, comments)
						Me.lblStatus.Text = txtKFNumCopies.Value.ToString & " copies ready to be imported by Kofax."
					End If
			End Select

			Me.ProgressBar.Value = 1
		Else

		End If
	End Sub

	Private Sub Reset_Kofax_Tab()
        'Reset the KofaxIt tab to default
        txtKFBatchName.Text = Now.ToString("MM-dd-yyyy HH:mm:ss")
		txtKFComments.Text = vbNullString
		cbKFBatchType.SelectedIndex = 9
		rbKFStandard.Checked = True
		btnKFImport.Enabled = True
		'Reset Account List
		gbKFAccounts.Visible = False
		lbKFAccounts.Items.Clear()
		If kfAccounts IsNot Nothing Then kfAccounts = New String() {}
		'Reset Copies
		lblKFCopies.Visible = False
		txtKFNumCopies.Visible = False
		txtKFNumCopies.Value = 1
	End Sub

#End Region

#Region "RightFax It Tab Region ----------------------------------------------------------------------------------"

	Private Sub btnRFax_Click(sender As Object, e As EventArgs) Handles btnRFax.Click
        Dim strServerName As String, strUsername As String, strPassword As String 'RightFax Server Strings
        Dim strRecName As String, strRecFax As String 'Fax Recipient Strings
        Dim bUseNTAuth As Boolean, bSaveRec As Boolean, bUseCoverSheet As Boolean
        Dim sFile As String
        Dim i As Integer = 0, faxesCount As Integer

        Dim objRightFax As FaxServer

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
            objRightFax = ConnectToServer(strServerName, strUsername, bUseNTAuth)
            objRightFax.OpenServer()
            faxesCount = dlgOpen.FileNames.Count
            ProgressBar.Maximum = faxesCount
            'Create and send each fax
            For Each sFile In dlgOpen.FileNames
                i += 1
                lblStatus.Text = String.Format("Faxing {0} of {1}.", i, faxesCount)
                Me.Refresh()
                Dim objFax As Fax = CreateFax(objRightFax, strRecName, strRecFax, sFile)
                SendFax(objFax)
                MoveFaxedFile(sFile)
                ProgressBar.Value += 1
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
        btnCAddContact.Enabled = False
        AddContactsToDatasetAccounts()

    End Sub

    Private Sub btnCGetAccounts_Click(sender As Object, e As EventArgs) Handles btnCGetAccounts.Click
        ResetContacts_Tab()
        btnCGetAccounts.Enabled = False

        'Get the list of accounts
        If (GetAccountsNeedingContacts(My.Settings.DatabaseFile)) Then
            'Fill the DataGridView
            Dim accCount As Integer = AddContact.contactDataSet.Tables("ContactsNeeded").Rows.Count
            If accCount > 0 Then
                lblContactStatus.Text = accCount.ToString & " Accounts Need A Contact"
                btnCAddContact.Enabled = True
                DataGridView.DataSource = AddContact.contactDataSet.Tables("ContactsNeeded")

                'Format Grid View
                DataGridView.Columns(0).Width = 90 'AccountNumber
                DataGridView.Columns(1).Width = 55 'DPA Type
                DataGridView.AutoResizeColumn(2) 'Sent To
                DataGridView.Columns(3).Width = 55 'Delivery Method
                DataGridView.Columns(4).Width = 22 'ContactAdded
                DataGridView.Columns(4).HeaderText = String.Empty
                DataGridView.Columns(4).Width = 30 'Key
            Else
                lblContactStatus.Text = "No Accounts Need A Contact"
            End If
        Else
            lblContactStatus.Text = "An error occurred."
        End If
        btnCGetAccounts.Enabled = True
    End Sub

    Private Sub btnStopAddContacts_Click(sender As Object, e As EventArgs) Handles btnStopAddContacts.Click
        _bStopContacts = True
    End Sub

    Private Sub ResetContacts_Tab()
        btnCAddContact.Enabled = False
        btnCGetAccounts.Enabled = True
        btnStopAddContacts.Enabled = False
        _bStopContacts = False
        DataGridView.DataSource = vbNull
        lblContactStatus.Text = String.Empty
    End Sub

    Private Sub Datagrid_MouseUp(sender As Object, e As MouseEventArgs) Handles DataGridView.MouseUp
        If (e.Button = MouseButtons.Right) Then
            Dim contactOptionsMenu As ContextMenuStrip = New ContextMenuStrip
            Dim hit As DataGridView.HitTestInfo = DataGridView.HitTest(e.X, e.Y)
            If hit.Type = DataGridViewHitTestType.Cell Then
                'Dim clickedCell As DataGridViewCell = DataGridView.Rows(hit.RowIndex).Cells(hit.ColumnIndex)
                DataGridView.ClearSelection()
                DataGridView.Rows(hit.RowIndex).Selected = True
                contactContextMenu.Show(DataGridView, e.Location)
            End If
        End If
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        Dim selectedRow As DataGridViewRow = DataGridView.SelectedRows.Item(0)
        Dim selectedItemKey As Integer = selectedRow.Cells("Key").Value
        DeleteContact(selectedItemKey)
        DataGridView.Rows.Remove(selectedRow)
    End Sub

#End Region

#Region "Convert Tab Region --------------------------------------------------------------------------------------"
    ''TODO Make convert tab use a background worker.
    Private Sub ConvertFiles(sFiles As String())
        Dim outputImage As String
        Dim outputExt As String
        Dim objWord As Microsoft.Office.Interop.Word.Application
        'Define the output folder and verify it exists.
        Dim sDestination As String = CONV_FOLDER
        CheckFolder(sDestination)

        'Disable tab items.
        Me.btnConvert.Enabled = False
        tabWordToTiff.AllowDrop = False


        'Setup progress bar
        Me.ProgressBar.Maximum = sFiles.Length
        Me.ProgressBar.Value = 0

        Try
            'Converts Word Documents or Images to .tif
            If (Me.rbConvertDOC.Checked) Then
                'Initiate word application object and minimize it
                objWord = CreateObject("Word.Application")
                objWord.WindowState = WdWindowState.wdWindowStateMinimize
                outputExt = "tiff"
            ElseIf (Me.rbConvertPDF.Checked) Then
                'Backup the PDF995 config files and create the pre-configed one.
                outputExt = "pdf"
                'Pdf995.BackupOriginalPdf995Ini()
                'Pdf995.CreatePdf995Ini(CONV_FOLDER)
                Throw New NotImplementedException()
                Exit Sub
            Else
                outputExt = "tif"
            End If

            'Print each file and report the progress to the form.
            For Each value In sFiles
                Me.lblStatus.Text = String.Format("Converting {0} of {1}", Me.ProgressBar.Value + 1, sFiles.Length)
                Me.Refresh()
                Dim fileName As String = Path.GetFileNameWithoutExtension(value)
                outputImage = [String].Format("{0}{1}.{2}", sDestination, fileName,
                                              outputExt)

                'Convert to tiff using the respective function.
                If (Me.rbConvertDOC.Checked) And (Me.rbConvertTiff.Checked) Then
                    DocToTiff(value, outputImage, objWord)
                ElseIf (Me.rbConvertIMAGE.Checked) And (Me.rbConvertTiff.Checked) Then
                    ImgToTiff(value, outputImage)
                ElseIf (Me.rbConvertIMAGE.Checked) And (Me.rbConvertPDF.Checked) Then
                    CreatePdf995Ini(CONV_FOLDER)
                    ImgToPdf(value, outputImage)
                End If
                Me.ProgressBar.Value += 1
            Next

            'Close word when all printing is done.
            If (Me.rbConvertDOC.Checked) Then
                objWord.Quit()
            End If
            objWord = Nothing

            Me.lblStatus.Text = "Conversion Complete!"
        Catch ex As System.Exception
            MessageBox.Show(ex.Message)
        End Try

        'Restore PDF995INI
        If (Me.rbConvertPDF.Checked) Then
            Pdf995.RestorePdf995Ini()
        End If
        'Re-enable tab
        Me.btnConvert.Enabled = True
        tabWordToTiff.AllowDrop = True
    End Sub

    Private Sub btnConvert_Click(sender As Object, e As EventArgs) Handles btnConvert.Click
        Reset()
        'If files are selected convert selection
        If dlgOpen.ShowDialog() = DialogResult.OK Then
            ConvertFiles(dlgOpen.FileNames)
        End If
    End Sub

    Private Sub _DragDrop(ByVal sender As Object, ByVal e As DragEventArgs) Handles tabWordToTiff.DragDrop
        'http://www.codemag.com/Article/0407031
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
            ConvertFiles(files)
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
			txtDPAdown.Text = vbNullString
			txtDPAmonthly.Text = vbNullString
			txtDPAmonthly.Enabled = True
        End If
    End Sub

    Private Sub btnBudgetBill_Click(sender As Object, e As EventArgs) Handles btnBudgetBill.Click
        Dim accNumber As String
        accNumber = mtxtDPAAcc.Text
        Call OpenCSSAcc(accNumber)
        Thread.Sleep(400)
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
        mtxtDPAAcc.Text = vbNullString
		txtDPAdown.Text = vbNullString
		txtDPAmonthly.Text = vbNullString
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
		Reset_Kofax_Tab()
        ResetContacts_Tab()
        Reset_TDriver_Tab()
    End Sub

    Private Sub Reset_ProgressBar()
        'Progress bar reset
        ProgressBar.Value = 0
        lblStatus.Text = ""
    End Sub

    Private Sub _DragEnter(ByVal sender As Object, ByVal e As DragEventArgs) _
        Handles tabWordToTiff.DragEnter, tabAddContact.DragEnter, tabTDrive.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
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

#Region "T:	Drive Tab Region -------------------------------------------------------------------------------------"

    Private Sub btnTEmails_Click(sender As Object, e As EventArgs) Handles btnTEmails.Click
        Reset_TDriver_Tab()
        CType(sender, Button).Enabled = False
        'btnTEmails.Enabled = False
        Reset_ProgressBar()
        If My.Settings.Physical_PrinterName = "DEFAULT" Then ChoosePhysicalPrinter()
        'Check if dialog was canceled
        If My.Settings.Physical_PrinterName = "DEFAULT" Then
            btnTEmails.Enabled = True
            Exit Sub
        End If

        Dim emailFolders() As String = {My.Settings.ActiveFolder, My.Settings.CutinFolder, My.Settings.AccInitFolder}
        For Each sFolder As String In emailFolders
            Try
                Dim files() As String = Directory.GetFiles(sFolder)
                ProcessFiles(files)
            Catch
                ''TODO Add invalid path handle.
            End Try
        Next
        CType(sender, Button).Enabled = True
    End Sub

    Private Sub DragDropTDrive(sender As Object, e As DragEventArgs) Handles tabTDrive.DragDrop
        Reset_ProgressBar()
        Me.lvTDriveFiles.Items.Clear()
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim files As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
            Try
                ProcessFiles(files)
            Catch ex As System.Exception
                MessageBox.Show(ex.Message)
                Return
            End Try
        End If
    End Sub


    Private Sub btnTDClear_Click(sender As Object, e As EventArgs)
        Me.lvTDriveFiles.Items.Clear()
    End Sub

    ''' <summary>
    ''' Open the file of the selected DPA.
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub TDriveListViewItemActivate_Open(sender As Object, e As EventArgs) _
        Handles lvTDriveFiles.ItemActivate
        'Cast the list view item tag as DPA
        Dim selectedDPA As DPA = CType(sender.SelectedItems(0).Tag, DPA)

        'Open the PDF if it was sent or the word file if it was not!
        If (selectedDPA.Sent) Then
            If (selectedDPA.DeliveryMethod <> DeliveryType.Mail) Then
                Process.Start(selectedDPA.FileToSend)
            End If
        Else
                Process.Start(selectedDPA.SourceFile)
        End If
    End Sub

    Private Sub Reset_TDriver_Tab()
        Me.btnTEmails.Enabled = True
        Me.lvTDriveFiles.Items.Clear()
    End Sub

	Private Sub ChoosePhysicalPrinter()
		Dim printDialog As New PrintDialog
		If printDialog.ShowDialog = System.Windows.Forms.DialogResult.OK Then
			Dim strPrinterName As String = printDialog.PrinterSettings.PrinterName
			My.Settings.Physical_PrinterName = strPrinterName
			My.Settings.Save()
		End If
	End Sub
#End Region

End Class
