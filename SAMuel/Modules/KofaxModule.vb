Imports System.IO

Module KofaxModule

    ''' <summary>
    ''' Creates an XML file which is read by Kofax Import Connector to create a batch
    ''' </summary>
    ''' <param name="docList">
    ''' List of per document tiff files stored in another list
    ''' </param>
    Sub CreateXML(ByRef docList As List(Of List(Of String)), batchName As String)
        Dim sPath As String
        Dim sFile As String
        Dim sList As List(Of String)
        Dim batchClass As String

        sPath = "C:\ACXMLAID\" 'Default Kofax Import Connector watch folder
        sFile = sPath & batchName & "-import.xml"
        batchClass = "eCorrespondence Fxd Pg"

        'Creates the directory if it doesn't exist
        If Not System.IO.Directory.Exists(sPath) Then
            System.IO.Directory.CreateDirectory(sPath)
        End If

        'Creates or Overwrites the XML file
        Dim fs1 As FileStream = New FileStream(sFile, FileMode.Create, FileAccess.Write)
        Dim s1 As StreamWriter = New StreamWriter(fs1)

        'XML Header
        s1.Write("<ImportSession>" & vbCrLf)
        s1.Write(vbTab & "<Batches>" & vbCrLf)
        s1.Write(vbTab & vbTab & "<Batch BatchClassName=""" & batchClass & """ EnableSingleDocProcessing = ""0"" Name=""" & batchName & """>" & vbCrLf)
        s1.Write(vbTab & vbTab & vbTab & "<Documents>" & vbCrLf)

        'Handle each document
        For Each sList In docList
            s1.Write(vbTab & vbTab & vbTab & vbTab & "<Document>" & vbCrLf)
            s1.Write(vbTab & vbTab & vbTab & vbTab & vbTab & "<Pages>" & vbCrLf)
            'Write each tiff file associated with the document
            For Each value In sList
                s1.Write(vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<Page ImportFileName=""" & value & """/>" & vbCrLf)
            Next
            s1.Write(vbTab & vbTab & vbTab & vbTab & vbTab & "</Pages>" & vbCrLf)
            s1.Write(vbTab & vbTab & vbTab & vbTab & "</Document>" & vbCrLf)
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
