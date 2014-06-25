Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Word = Microsoft.Office.Interop.Word
Imports System.IO
Imports System.Drawing.Imaging


Public Class frmMain
    'Form wide variables
    Dim bNextPressed As Boolean
    Dim bRejectPressed As Boolean
    Dim bCancelPressed As Boolean
    Dim bAuditMode As Boolean

    Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click
        Dim oApp As Outlook.Application = New Outlook.Application
        Dim sDestination As String
        Dim sFile As String, sFileExt As String
        Dim outTiff As String
        Dim attachmentImg As Image, modifiedImg As Image = Nothing
        Dim ms As MemoryStream = Nothing
        Dim docList As New List(Of List(Of String))
        Dim docTiffList As New List(Of String)
        Dim i As Integer = 0, emailCount As Integer
        Dim samEmails As New List(Of SAM_Email)
        Dim sEmail As SAM_Email
        Dim rand As New Random


        Reset_Outlook_Tab()
        clbSelectedEmails.Items.Clear()
        Reset_ProgressBar()

        'enable buttons
        btnCancel.Enabled = True
        btnReject.Enabled = True
        btnNext.Visible = True
        btnNext.Enabled = True
        btnRun.Enabled = False
        btnRun.Visible = False

        'Create each SAM Email
        Try
            For Each value In oApp.ActiveExplorer.Selection
                Try
                    lblStatus.Text = "Saving attachments..."
                    Me.Refresh()
                    sEmail = New SAM_Email(value)
                    sEmail.Regex()
                    samEmails.Add(sEmail)
                    clbSelectedEmails.Items.Add("[" & sEmail.AttachmentsCount.ToString & "] " & sEmail.Subject)
                Catch ex As Exception
                    'Move to next email
                    Continue For
                End Try
            Next
        Catch ex As Exception
            MsgBox("No emails within Outlook selected.", MsgBoxStyle.Exclamation)
            Exit Sub
        End Try

        emailCount = samEmails.Count
        ProgressBar.Maximum = emailCount

        oApp = Nothing

        'Process each email
        For Each sEmail In samEmails
            Try
                'Force an audit if the email is found to be not valid.
                If (Not sEmail.IsValid) Then
                    bAuditMode = True
                    lblOutlookMessage.Text = "Invalid Email Found!"
                End If

                'Process each attachment within the email
                If sEmail.AttachmentsCount > 0 Then
                    docTiffList = New List(Of String)
                    For Each sFile In sEmail.Attachments
                        'Verify a valid attachment file type.
                        sFileExt = EmailProcessing.ValidateAttachmentType(sFile)
                        If (sFileExt Is vbNullString) Then
                            'Delete the file and move to next attachment
                            System.IO.File.Delete(sFile)
                            Continue For
                        ElseIf sFileExt = ".pdf" Then
                            EmailProcessing.ParsePDFImgs(sFile)
                            Continue For
                        End If
                        'Load the attachment
                        attachmentImg = Image.FromFile(sFile)
                        picImage.Image = attachmentImg
                        'Display email info when in auditmode
                        If (bAuditMode) Then
                            txtAcc.Text = sEmail.Account
                            txtSubject.Text = sEmail.Subject
                            txtFrom.Text = sEmail.From
                        End If

                        'Wait for user validation of attachment
                        Do Until (bNextPressed = True Or bRejectPressed = True Or bCancelPressed = True Or bAuditMode = False)
                            Application.DoEvents()
                        Loop

                        'Update sEmail account if its in audit mode
                        If (bAuditMode) Then
                            sEmail.Account = txtAcc.Text
                        End If

                        bAuditMode = chkAuditMode.Checked

                        If bCancelPressed Then
                            'When canceled, release image and reset form and end the routine
                            picImage.Image = Nothing
                            attachmentImg.Dispose()
                            attachmentImg = Nothing
                            Me.Dispose()
                            Reset_Outlook_Tab()
                            Exit Sub
                        ElseIf bRejectPressed Then
                            '------ LOG reject action? ---------
                        ElseIf bNextPressed Or bAuditMode = False Then
                            i += 1
                            'Resize the image if it is too large
                                lblStatus.Text = "Resizing Image..."
                                Me.Refresh()
                                ms = ImageProcessing.ResizeImage(attachmentImg)
                                modifiedImg = Image.FromStream(ms)
                                Application.DoEvents()


                                'Release original attachment so it can be removed when done processing
                                picImage.Image = Nothing
                                attachmentImg.Dispose()
                                attachmentImg = Nothing

                                'Show the current image as it's being modified
                                picImage.Image = modifiedImg

                                'Add account number as a watermark
                                lblStatus.Text = "Adding Watermark..."
                                Me.Refresh()
                                EmailProcessing.Add_Watermark(modifiedImg, sEmail.Account) ''add suffix handing
                                'Convert the image to Grayscale
                                lblStatus.Text = "Converting to Grayscale..."
                                Me.Refresh()
                                EmailProcessing.MakeGrayscale(modifiedImg)
                                'Save the edited attachment as tiff and add to list
                                lblStatus.Text = "Converting to Tiff..."
                                Me.Refresh()
                                sDestination = My.Settings.savePath + "tiffs\" + sEmail.From + "\"
                                GlobalModule.CheckFolder(sDestination)
                                outTiff = [String].Format("{0}{1}_{2}.tiff", sDestination, rand.Next(10000).ToString, sEmail.Account)
                                ImageProcessing.CompressTiff(modifiedImg, outTiff)
                                docTiffList.Add(outTiff)
                                Me.Refresh()
                        Else
                            'Undesired state. Log this
                            LogAction(99, "Outlook buttonState: " & bCancelPressed.ToString & "," & bRejectPressed.ToString & "," & bNextPressed.ToString & "," & bAuditMode.ToString)
                        End If

                        'Reset variables
                        bNextPressed = False
                        bRejectPressed = False

                        'Release image
                        picImage.Image = Nothing
                        If Not IsNothing(modifiedImg) Then modifiedImg.Dispose()
                        attachmentImg = Nothing
                        If Not IsNothing(ms) Then ms.Dispose()
                        ms = Nothing

                        Application.DoEvents()
                        Me.Refresh()

                        'Delete the saved email attachment
                        System.IO.File.Delete(sFile)
                        lblStatus.Text = ""
                        'Checks the email in the check list box
                        clbSelectedEmails.SetItemChecked(ProgressBar.Value, True)

                    Next
                    'Add document to list of documents with tiffs
                    docList.Add(docTiffList)

                Else
                    'No attachments
                    LogAction(50, String.Format("{0} - {1} was skipped. No attachments", sEmail.Subject, sEmail.From))
                End If

                bNextPressed = False
                ProgressBar.Value += 1
                Me.Refresh()
                Application.DoEvents()
            Catch ex As Exception
                'Unknown error with email.
                LogAction(98, String.Format("{0} - {1} : {2}", sEmail.Subject, sEmail.From, ex.ToString))
            End Try
        Next

        ''**Handle No emails complete **''

        'Create XML if user decides to.
        If MsgBox("Email processing complete. Continue to create XML?", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
            lblStatus.Text = "Creating XML File..."
            Me.Refresh()
            KofaxModule.CreateXML(docList, "SAMuel - " & docList.Count.ToString & " emails.")
        End If

        lblStatus.Text = "DONE!"

        Reset_Outlook_Tab()

    End Sub

    Private Sub btnConvert_Click(sender As Object, e As EventArgs) Handles btnConvert.Click
        'Converts Word Documents to .tif using MODI
        Reset()
        'If files are selected continue code
        If dlgOpen.ShowDialog() = DialogResult.OK Then
            Try
                ConvertToTiff.WordDocs(dlgOpen.FileNames)

            Catch ex As Exception
                MessageBox.Show(ex.Message)
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
        clbSelectedEmails.Items.Clear()
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

        'clear text fields
        rtbEmailBody.Text = ""
        txtAcc.Text = ""
        txtFrom.Text = ""
        txtAcc.Text = ""
        txtSubject.Text = ""
        lblOutlookMessage.Text = ""
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        bCancelPressed = True
    End Sub

    Private Sub btnNext_Click_1(sender As Object, e As EventArgs) Handles btnNext.Click
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
                ConvertToTiff.WordDocs(files)

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
        batchName = Me.txtKFBatchName.Text
        dlgOpen.Filter = "TIFF Images|*.tif;*.tiff"
        If dlgOpen.ShowDialog() = DialogResult.OK Then
            'Disabled until handle two different inputs
            'KofaxModule.CreateXML((dlgOpen.FileNames).ToList, batchName)
        End If

    End Sub
End Class