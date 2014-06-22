Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Word = Microsoft.Office.Interop.Word
Imports System.IO
Imports System.Drawing.Imaging


Public Class frmMain
    'Form wide variables
    Dim bNextPressed As Boolean
    Dim bRejectPressed As Boolean
    Dim bCancelPressed As Boolean

    Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click
        Dim oApp As Outlook.Application = New Outlook.Application
        Dim oMsg As Outlook.MailItem
        Dim oAtt As Outlook.Attachment
        Dim oSelection As Outlook.Selection
        Dim sDestination As String = My.Settings.savePath
        Dim sFile As String, sFileExt As String, editedImg As String
        Dim outTiff As String
        Dim strREGEXed As String, strSubject As String
        Dim attachmentImg As Image
        Dim tiffList As New List(Of String)
        Dim i As Integer = 0, emailCount As Integer

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

        '**TODO** Add no selection handling
        oSelection = oApp.ActiveExplorer.Selection

        For Each oMsg In oSelection
            clbSelectedEmails.Items.Add("[" & oMsg.Attachments.Count.ToString & "] " & oMsg.Subject)
        Next

        emailCount = oSelection.Count

        ProgressBar.Maximum = emailCount
        'Process each email selected
        For Each oMsg In oSelection
            'Parse data from the email
            strSubject = oMsg.Subject
            txtSubject.Text = strSubject
            txtFrom.Text = oMsg.SenderName
            rtbEmailBody.Text = oMsg.Body

            'REGEX subject line for Account number or customer #
            If strSubject IsNot vbNullString Then
                strREGEXed = GlobalModule.RegexAccount(strSubject)
                If strREGEXed <> "ACC# NOT FOUND" Then
                    txtAcc.Text = strREGEXed
                Else
                    'Look for Customer Number
                    strREGEXed = GlobalModule.RegexCustomer(strSubject)
                    If strREGEXed <> "CUST# NOT FOUND" Then
                        txtAcc.Text = strREGEXed
                    Else
                        strREGEXed = "REGEX_BODY"
                    End If
                End If
            Else
                strREGEXed = "REGEX_BODY"
            End If

            'REGEX Body for Account number or customer #
            If strREGEXed = "REGEX_BODY" Then
                If rtbEmailBody.Text IsNot vbNullString Then
                    strREGEXed = GlobalModule.RegexAccount(rtbEmailBody.Text)
                    If strREGEXed <> "ACC# NOT FOUND" Then
                        txtAcc.Text = strREGEXed
                    Else
                        strREGEXed = GlobalModule.RegexCustomer(rtbEmailBody.Text)
                        If strREGEXed <> "CUST# NOT FOUND" Then
                            txtAcc.Text = strREGEXed
                        Else
                            txtAcc.Text = "ACCOUNT UNKNOWN"
                            chkAuditMode.Checked = True
                        End If
                    End If

                Else
                    txtAcc.Text = "ACCOUNT UNKNOWN"
                    chkAuditMode.Checked = True
                End If
            End If

            'Process each attachment within the email
            If oMsg.Attachments.Count > 0 Then
                For Each oAtt In oMsg.Attachments
                    sFile = sDestination & oAtt.FileName
                    'Verify a valid attachment file type
                    sFileExt = Path.GetExtension(sFile).ToLower
                    If sFileExt = ".tiff" Or sFileExt = ".png" Or _
                            sFileExt = ".jpg" Or sFileExt = ".jpeg" Or _
                            sFileExt = ".tif" Or sFileExt = ".gif" Or _
                            sFileExt = ".bmp" Then
                        'Save the attachment then load the preview
                        oAtt.SaveAsFile(sFile)
                        attachmentImg = Drawing.Image.FromFile(sFile)
                        picImage.Image = attachmentImg

                        'Wait for user validation of attachment
                        Do Until (bNextPressed = True Or bRejectPressed = True Or bCancelPressed = True Or Me.chkAuditMode.Checked = False)
                            Application.DoEvents()
                        Loop
                        If bCancelPressed Then
                            'When canceled, reset form and end the routine
                            'Release image
                            picImage.Image.Dispose()
                            picImage.Image = Nothing

                            attachmentImg.Dispose()
                            attachmentImg = Nothing

                            Reset_Outlook_Tab()
                            Exit Sub
                        ElseIf bNextPressed Or Me.chkAuditMode.Checked = False Then
                            i += 1
                            'Add account number as a watermark
                            lblStatus.Text = "Adding Watermark..."
                            Me.Refresh()
                            'EmailProcessing.ReSize_IMG(attachmentImg)
                            'EmailProcessing.Resize_Image(attachmentImg)
                            EmailProcessing.Add_Watermark(attachmentImg, txtAcc.Text) ''add suffix handing
                            lblStatus.Text = "Converting to Tiff..."
                            Me.Refresh()
                            'Save the edited attachment as tiff and add to list
                            outTiff = [String].Format("{0}{1}{2}.tiff", sDestination, i & "_", txtAcc.Text)
                            picImage.Image.Save(outTiff, System.Drawing.Imaging.ImageFormat.Tiff)
                            tiffList.Add(outTiff)
                            Me.Refresh()
                        ElseIf bRejectPressed Then
                            '------ LOG reject action? ---------
                        End If

                        'Reset variables
                        bNextPressed = False
                        bRejectPressed = False

                        'Release image
                        picImage.Image.Dispose()
                        picImage.Image = Nothing

                        attachmentImg.Dispose()
                        attachmentImg = Nothing

                        Application.DoEvents()
                        Me.Refresh()
                        'Delete the saved email attachment
                        ''System.IO.Directory.GetFiles() ' Get all files within a folder ** useful later **
                        System.IO.File.Delete(sFile)
                        lblStatus.Text = ""
                        'Checks the email in the check list box
                        clbSelectedEmails.SetItemChecked(ProgressBar.Value, True)
                    Else
                        'Invalid Email Attachment

                    End If
                Next
            Else
                'No attachments

            End If

            bNextPressed = False
            ProgressBar.Value += 1
            Me.Refresh()
            Application.DoEvents()
        Next

        lblStatus.Text = "Creating XML File..."
        Me.Refresh()
        KofaxModule.CreateXML(tiffList, "SAMuel - " & emailCount & " emails")
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
                Return
            End Try
        End If

        'Cleanup
        btnConvert.Enabled = True
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnReject.Click
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
        If (My.Settings.savePath = "" Or Not System.IO.Directory.Exists(My.Settings.savePath)) Then
            Dim defaultPath As String = Environment.GetEnvironmentVariable("userprofile") & "\SAMuel\"
            'Create the folder if it does not exist
            If (Not System.IO.Directory.Exists(defaultPath)) Then
                System.IO.Directory.CreateDirectory(defaultPath)
            End If
            My.Settings.savePath = defaultPath
            My.Settings.Save()
        End If
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

        Dim objRightFax As RFCOMAPILib.FaxServer
        Dim objFax As RFCOMAPILib.Fax

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
            'Create and send each fax
            For Each sFile In dlgOpen.FileNames
                objFax = RightFax.CreateFax(objRightFax, strRecName, strRecFax, sFile)
                RightFax.SendFax(objFax)
                RightFax.MoveFaxedFile(sFile)
                objFax = Nothing
            Next
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
            KofaxModule.CreateXML((dlgOpen.FileNames).ToList, batchName)
        End If

    End Sub
End Class