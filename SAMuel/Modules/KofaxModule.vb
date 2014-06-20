Imports System.IO

Module KofaxModule

    ''' <summary>
    ''' Creates an XML file which is read by Kofax Import Connector to create a batch
    ''' </summary>
    ''' <param name="fileList">
    ''' Array of files to be added to the XML file
    ''' </param>
    Sub CreateXML(fileList() As String)
        Dim sPath As String
        Dim batchName As String
        Dim batchClass As String

        sPath = "C:\ACXMLAID"
        batchName = frmMain.txtKFBatchName.Text
        batchClass = "eCorrespondence Fxd Pg"
        If Not System.IO.Directory.Exists(sPath) Then
            System.IO.Directory.CreateDirectory(sPath)
        End If

        'check the file
        Dim fs As FileStream = New FileStream(sPath & "\" & batchName & "-import.xml", FileMode.OpenOrCreate, FileAccess.ReadWrite)
        Dim s As StreamWriter = New StreamWriter(fs)
        s.Close()
        fs.Close()

        Dim fs1 As FileStream = New FileStream(sPath & "\" & batchName & "-import.xml", FileMode.Append, FileAccess.Write)
        Dim s1 As StreamWriter = New StreamWriter(fs1)

        'XML Header
        s1.Write("<ImportSession>" & vbCrLf)
        s1.Write(vbTab & "<Batches>" & vbCrLf)
        s1.Write(vbTab & vbTab & "<Batch BatchClassName=""" & batchClass & """ EnableSingleDocProcessing = ""0"" NAME=""" & batchName & """>" & vbCrLf)
        s1.Write(vbTab & vbTab & vbTab & "<Documents>" & vbCrLf)

        'Handle each file
        For Each value In fileList
            s1.Write(vbTab & vbTab & vbTab & vbTab & "<Document>" & vbCrLf)
            s1.Write(vbTab & vbTab & vbTab & vbTab & vbTab & "<Pages>" & vbCrLf)
            s1.Write(vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<Page ImportFileName=""" & value & """/>" & vbCrLf)
            s1.Write(vbTab & vbTab & vbTab & vbTab & vbTab & "</Pages>" & vbCrLf)
            s1.Write(vbTab & vbTab & vbTab & vbTab & "</Document>" & vbCrLf)
            'Me.Refresh()
        Next

        'End XML File
        s1.Write(vbTab & vbTab & vbTab & "</Documents>" & vbCrLf)
        s1.Write(vbTab & vbTab & "</Batch>" & vbCrLf)
        s1.Write(vbTab & "</Batches>" & vbCrLf)
        s1.Write("</ImportSession>" & vbCrLf)
        s1.Close()
        fs1.Close()
    End Sub

End Module
