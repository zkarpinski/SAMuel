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
        Dim sFile As String, sFileExt As String, sEditedImg As String
        Dim attachmentImg As Image

        Reset_Outlook_Tab()
        Reset_ProgressBar()

        'enable buttons
        btnCancel.Enabled = True
        btnReject.Enabled = True
        btnNext.Visible = True
        btnNext.Enabled = True
        btnRun.Enabled = False
        btnRun.Visible = False

        oSelection = oApp.ActiveExplorer.Selection
        ProgressBar.Maximum = oSelection.Count
        'Process each email selected
        For Each oMsg In oSelection
            'Parse data from the email
            txtSubject.Text = oMsg.Subject
            txtFrom.Text = oMsg.SenderName
            txtAcc.Text = regexAccount(oMsg.Subject)
            rtbEmailBody.Text = oMsg.Body

            'Process each attachment within the email
            If oMsg.Attachments.Count > 0 Then
                For Each oAtt In oMsg.Attachments
                    sFile = sDestination & oAtt.FileName
                    'Verify a valid attachment file type
                    sFileExt = Path.GetExtension(sFile).ToLower
                    If sFileExt = ".tiff" Or sFileExt = ".png" Or _
                            sFileExt = ".jpg" Or sFileExt = ".jpeg" Or _
                            sFileExt = ".tif" Or sFileExt = ".gif" Then
                        'Save the attachment then load the preview
                        oAtt.SaveAsFile(sFile)
                        attachmentImg = Drawing.Image.FromFile(sFile)
                        'picImage.Image = New Bitmap(sFile)
                        picImage.Image = attachmentImg
                        'Wait for user validation of attachment
                        Do Until (bNextPressed = True Or bRejectPressed = True Or bCancelPressed = True)
                            Application.DoEvents()
                        Loop
                        If bCancelPressed Then
                            'When canceled, reset form and end the routine
                            Reset_Outlook_Tab()
                            Exit Sub
                        ElseIf bNextPressed Then
                            'Add account number as a watermark
                            lblStatus.Text = "Adding Watermark..."
                            sEditedImg = Add_Watermark(attachmentImg, txtAcc.Text) ''add suffix handing
                            lblStatus.Text = "Converting to Tiff..."
                            Convert_Image_To_Tif(sEditedImg, False)
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
                        'Delete the saved email attachment
                        ''System.IO.Directory.GetFiles() ' Get all files within a folder ** useful later **
                        System.IO.File.Delete(sFile)
                        lblStatus.Text = ""
                    End If
                Next
            End If
            'Wait for user to move to next email
            'Do Until (bNextPressed = True)
            '   Application.DoEvents()
            'Loop
            bNextPressed = False
            ProgressBar.Value += 1
        Next
        lblStatus.Text = "DONE!"

        Reset_Outlook_Tab()

    End Sub

    Private Sub btnConvert_Click(sender As Object, e As EventArgs) Handles btnConvert.Click
        'Converts Word Documents to .tif using MODI
        Dim objWdDoc As Word.Document
        Dim objWord As Word.Application
        Dim sDoc As String
        Dim sDestination As String = My.Settings.savePath
        Dim sFileName As String

        Reset()

        'If files are selected continue code
        If dlgOpen.ShowDialog() = DialogResult.OK Then
            'Initiate word application object and minimize it
            objWord = CreateObject("Word.Application")
            objWord.WindowState = Word.WdWindowState.wdWindowStateMinimize
            'Set active printer to Fax
            objWord.ActivePrinter = "Microsoft Office Document Image Writer"
            'Determine progressbar maximum value from number of files chosen
            ProgressBar.Maximum = dlgOpen.FileNames.Count()
            For Each sDoc In dlgOpen.FileNames
                'Get the file name
                sFileName = Path.GetFileNameWithoutExtension(sDoc)
                'Open the document within word
                objWdDoc = objWord.Documents.Open(sDoc)
                objWord.Visible = False
                'Print to Tiff
                objWdDoc.PrintOut(PrintToFile:=True, OutputFileName:=sDestination & sFileName & ".tif")
                'Release document
                objWdDoc.Close()
                'Progress the progress bar
                ProgressBar.Value += 1
            Next
            lblStatus.Text = "DONE!"
            objWord.Quit()
        End If

        'Cleanup
        btnConvert.Enabled = True
        objWdDoc = Nothing
        objWord = Nothing
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

        'If files are selected continue code
        If dlgOpen.ShowDialog() = DialogResult.OK Then
            ''Do RightFax stuff here
        End If
    End Sub

    Private Function Add_Watermark(ByVal img As Image, ByVal accountNumber As String, Optional ByVal suffix As String = vbNullString) As String
        'Adds the account number as a watermark to the image
        'http://bytescout.com/products/developer/watermarkingsdk/how_to_add_text_watermark_to_image_using_dotnet_in_vb.html
        Dim graphics As Graphics
        Dim font As Font
        Dim brush As SolidBrush
        Dim point As PointF
        Dim savedFile As String

        'Determine if file needs a suffix
        If suffix = vbNullString Then
            savedFile = My.Settings.savePath + accountNumber + ".jpg"
        Else
            savedFile = My.Settings.savePath + accountNumber + "_" + suffix + ".jpg"
        End If

        ' Account Number Watermark Settings
        font = New Font("Times New Roman", 30.0F)
        point = New PointF(500, 50) '(x,y)
        brush = New SolidBrush(Color.FromArgb(150, Color.Black))

        ' Create graphics from image
        graphics = Drawing.Graphics.FromImage(img)
        ' Draw the Account Number Watermark
        graphics.DrawString(accountNumber, font, brush, point)
        ' Save image
        img.Save(savedFile)
        'Clear Resources
        graphics.Dispose()
        graphics = Nothing
        img.Dispose()
        img = Nothing
        Return savedFile
    End Function

    ''**NOTE**This function needs an overhaul
    Private Sub Convert_Image_To_Tif(fileNames As String, isMultiPage As Boolean)
        'Converts the image to a tiff
        'http://code.msdn.microsoft.com/windowsdesktop/VBTiffImageConverter-f8fdfd7f/sourcecode?fileId=26346&pathId=565080007
        Using encoderParams As New EncoderParameters(1)
            Dim tiffCodecInfo As ImageCodecInfo = ImageCodecInfo.GetImageEncoders(). _
                   First(Function(ie) ie.MimeType = "image/tiff")

            Dim tiffPaths As String() = Nothing
            If isMultiPage Then
                tiffPaths = New String(1) {}
                Dim tiffImg As Image = Nothing
                Try
                    Dim i As Integer = 0
                    While i < fileNames.Length
                        If i = 0 Then
                            tiffPaths(i) = [String].Format("{0}\{1}.tiff",
                                        Path.GetDirectoryName(fileNames(i)),
                                        Path.GetFileNameWithoutExtension(fileNames(i)))

                            ' Initialize the first frame of multi-page tiff. 
                            tiffImg = Image.FromFile(fileNames(i))
                            encoderParams.Param(0) = New EncoderParameter(
                                Encoder.SaveFlag, DirectCast(EncoderValue.MultiFrame, 
                                EncoderParameterValueType))
                            tiffImg.Save(tiffPaths(i), tiffCodecInfo, encoderParams)
                        Else
                            ' Add additional frames. 
                            encoderParams.Param(0) = New EncoderParameter(
                                Encoder.SaveFlag,
                                DirectCast(EncoderValue.FrameDimensionPage, 
                                EncoderParameterValueType))
                            Using frame As Image = Image.FromFile(fileNames(i))
                                tiffImg.SaveAdd(frame, encoderParams)
                            End Using
                        End If

                        If i = fileNames.Length - 1 Then
                            ' When it is the last frame, flush the resources and closing. 
                            encoderParams.Param(0) = New EncoderParameter(
                                Encoder.SaveFlag,
                                DirectCast(EncoderValue.Flush, EncoderParameterValueType))
                            tiffImg.SaveAdd(encoderParams)
                        End If
                        System.Math.Max(System.Threading.Interlocked.Increment(i), i - 1)
                    End While
                Finally
                    If tiffImg IsNot Nothing Then
                        tiffImg.Dispose()
                        tiffImg = Nothing
                    End If
                End Try
            Else
                tiffPaths = New String(fileNames.Length) {}

                Dim i As Integer = 0
                While i < fileNames.Length
                    tiffPaths(i) = [String].Format("{0}{1}.tiff",
                                       My.Settings.savePath,
                                       "testImage")

                    ' Save as individual tiff files. 
                    Using tiffImg As Image = Image.FromFile(fileNames)
                        tiffImg.Save(tiffPaths(i), ImageFormat.Tiff)
                    End Using
                    System.Math.Max(System.Threading.Interlocked.Increment(i), i - 1)
                End While
            End If
        End Using
    End Sub

    Private Sub OptionsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OptionsToolStripMenuItem.Click
        frmOptions.Show()
    End Sub

    Private Sub Convert_Image_To_Tif(sEditedImg As String, Optional p2 As Object = Nothing, Optional p3 As Boolean = Nothing)
        Throw New NotImplementedException
    End Sub

End Class