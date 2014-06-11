Attribute VB_Name = "RightFax_It"
'RightFax It
' VBA macro that automatically faxes the selected files, if they are in the correct format.
' Created by: Zachary Karpinski

'References:
'   Microsoft Word
'   RightFax COM API
'   RightFax Object Type

Sub RightFax_It()
Dim intFd As Integer
Dim i As Integer
Dim strFile As String
Dim strSplit As Variant
Dim objWord As Word.Application
Dim fd As FileDialog
Dim strName As String, strFax As String, strAcc As String
Dim newFax As RFCOMAPILib.Fax
Dim FaxAPI As RFCOMAPILib.FaxServer
Set FaxAPI = New RFCOMAPILib.FaxServer
Set objWord = New Word.Application

    'Setup Open file dialog using Word
Set fd = objWord.FileDialog(msoFileDialogFilePicker)
fd.AllowMultiSelect = True
    fd.InitialFileName = Environ$("USERPROFILE") & "\My Documents\"


'Connect to RightFax Server
    FaxAPI.ServerName = "SERVER_NAME"
FaxAPI.Protocol = 4
    FaxAPI.AuthorizationUserID = "USERNAME"
FaxAPI.UseNTAuthentication = False
FaxAPI.OpenServer

intFd = fd.Show
objWord.WindowState = wdWindowStateMinimize
If intFd <> 0 Then
        For i = 1 To fd.SelectedItems.count
            strFile = fd.SelectedItems(i)
            'Parse receipant name, fax and account
            'Ideal file name: Fax-12345-67890-To-NAME_HERE-Number-1-888-888-8888.doc
            strSplit = Split(strFile, "-")
            strName = Replace(strSplit(4), "_", " ") 'Name: Remove _ also
            strFax = strSplit(6) + strSplit(7) + strSplit(8) + strSplit(9) 'Fax Number
            strFax = Replace(strFax, ".doc", "") 'Remove .doc if no attn
            strAcc = strSplit(1) + "-" + strSplit(2) 'Acc Number

            ''TODO: Validate the fax number and name

            'Create individual fax and send
            newFax = FaxAPI.CreateObject(RFCOMAPILib.CreateObjectType.coFax)
            newFax.ToName = strName
            newFax.ToFaxNumber = strFax
            newFax.Attachments.Add(strFile, False)
            newFax.Send()

            'Log the action
            LogAction("Faxed document to " + strName + " Account:" + strAcc)
            'clear fax obj
            newFax = Nothing
        Next i
    
    'Close word
    objWord.Quit
    Set objWord = Nothing
    
        MsgBox "SENT"
Else
    'Something here?
End If
FaxAPI.CloseServer
End Sub


Private Sub LogAction(action As String)
'Writes a logfile with the input string
'http://stackoverflow.com/questions/10403517/how-to-make-an-external-log-using-excel-vba
    Dim logFile As String
    Dim nFileNum As Long
    logFile = Environ$("USERPROFILE") & "\My Documents\RightFax_It.log"
    action = Now & " " + action

    nFileNum = FreeFile
    Open logFile For Append As #nFileNum   ' create the file if it doesn't exist
    Print #nFileNum, action                ' append information
    Close #nFileNum
End Sub

