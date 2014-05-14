Imports Outlook = Microsoft.Office.Interop.Outlook
Imports Word = Microsoft.Office.Interop.Word
Public Class Form1

    Dim btnNextPressed As Boolean

    Private Sub btnRun_Click(sender As Object, e As EventArgs) Handles btnRun.Click
        Dim oApp As Outlook.Application = New Outlook.Application
        Dim objFolder As Outlook.MAPIFolder = oApp.Session.Folders.GetFirst
        objFolder = objFolder.Folders.Item("Inbox")
        Dim oItems As Outlook.Items = objFolder.Items
        Dim oMsg As Outlook.MailItem
        Dim oAtt As Outlook.Attachment
        Dim sFile As String

        For Each oMsg In oItems
            Debug.Print(oMsg.Subject)
            txtSubject.Text = oMsg.Subject
            txtFrom.Text = oMsg.SenderName
            Debug.Print(oMsg.Attachments.Count)
            If oMsg.Attachments.Count > 0 Then
                For Each oAtt In oMsg.Attachments
                    sFile = Environment.GetEnvironmentVariable("userprofile") & "\Desktop\" & oAtt.FileName
                    oAtt.SaveAsFile(sFile)
                    picImage.Image = New Bitmap(sFile)
                    Do Until (btnNextPressed = True)
                        Application.DoEvents()
                    Loop
                    btnNextPressed = False

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
        Dim sDesktop As String = _
            Environment.GetEnvironmentVariable("userprofile") & "\Desktop\"
        If dlgOpen.ShowDialog() = DialogResult.OK Then
            Dim i As Integer
            i = 0
            objWord = CreateObject("Word.Application")
            objWord.ActivePrinter = "Fax"

            For Each sDoc In dlgOpen.FileNames
                i += 1
                objWdDoc = objWord.Documents.Open(sDoc)
                objWord.Visible = False
                'Print to Tiff
                objWdDoc.PrintOut(PrintToFile:=True, OutputFileName:=sDesktop & i & ".tiff")
                objWdDoc.Close()
            Next
            objWord.Quit()
        End If

        'Range:=WdPrintOutRange.wdPrintAllDocument, _
        '                     OutputFileName:=sDesktop & "test.tiff", _
        '                      Item:=WdPrintOutItem.wdPrintDocumentContent, _
        '                      PrintToFile:=True)

        'General Cleanup
        objWdDoc = Nothing
        objWord = Nothing
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        btnNextPressed = True
    End Sub
End Class
