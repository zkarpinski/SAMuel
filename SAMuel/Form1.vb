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

        For Each oMsg In oItems
            Debug.Print(oMsg.Subject)
            txtSubject.Text = oMsg.Subject
            txtFrom.Text = oMsg.SenderName
            Debug.Print(oMsg.Attachments.Count)
            If oMsg.Attachments.Count > 0 Then
                For Each oAtt In oMsg.Attachments
                    picImage.Image = New Bitmap(oAtt.PathName)
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
        Dim sDesktop As String = _
            Environment.GetEnvironmentVariable("userprofile") & "\Desktop\"

        objWord = CreateObject("Word.Application")
        objWdDoc = objWord.Documents.Open(sDesktop & "testdocument.docx")
        objWord.Visible = True

        'Select Printer
        objWord.ActivePrinter = "Fax"

        'Print to Tiff
        objWdDoc.PrintOut(PrintToFile:=True, OutputFileName:=sDesktop & "test.tiff")
        'Range:=WdPrintOutRange.wdPrintAllDocument, _
        '                     OutputFileName:=sDesktop & "test.tiff", _
        '                      Item:=WdPrintOutItem.wdPrintDocumentContent, _
        '                      PrintToFile:=True)
        'Close Document
        objWdDoc.Close()
        'Close Word
        objWord.Quit()
        'General Cleanup
        objWdDoc = Nothing
        objWord = Nothing
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        btnNextPressed = True
    End Sub
End Class
