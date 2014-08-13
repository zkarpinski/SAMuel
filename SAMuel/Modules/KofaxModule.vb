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

    Sub CreateXML(ByRef docList As List(Of String), batchName As String, corrType As String, corrSource As String, Optional strComments As String = vbNullString)
        Dim sPath As String
        Dim sFile As String
        Dim sAccountNumber As String
        Dim _IsCustomerNumber As Boolean
        Dim batchClass As String

        sPath = "C:\ACXMLAID\" 'Default Kofax Import Connector watch folder
        sFile = sPath & batchName & "-import.xml"
        batchClass = "eCorrespondence Fax"

        'Creates the directory if it doesn't exist
        If Not System.IO.Directory.Exists(sPath) Then
            System.IO.Directory.CreateDirectory(sPath)
        End If

        'Creates or Overwrites the XML file
        Dim fs1 As FileStream = New FileStream(sFile, FileMode.Create, FileAccess.Write)
        Dim s1 As StreamWriter = New StreamWriter(fs1)

        'XML Header
        s1.Write("<ImportSession>")
        s1.Write(vbTab & "<Batches>")
        s1.Write(vbTab & vbTab & "<Batch BatchClassName=""" & batchClass & """ EnableSingleDocProcessing = ""0"" Name=""" & batchName & """>")
        s1.Write(vbTab & vbTab & vbTab & "<Documents>")



        'Write each tiff file as a document
        For Each value In docList

            _IsCustomerNumber = False

            sAccountNumber = RegexAcc(value, "\d{5}-\d{5}")
            If sAccountNumber = "X" Then
                sAccountNumber = RegexAcc(value, "\d{10}")
                If sAccountNumber = "X" Then
                    sAccountNumber = RegexAcc(value, "\d{9}")
                    If sAccountNumber = "X" Then
                        sAccountNumber = ""
                    Else
                        _IsCustomerNumber = True
                    End If
                End If
            End If

            'Remove Hyphen from account number
            sAccountNumber = sAccountNumber.Replace("-", "")

            s1.Write(vbTab & vbTab & vbTab & vbTab & "<Document FormType = ""eCorrespondence Fax"">")
            s1.Write(vbTab & vbTab & vbTab & vbTab & vbTab & "<IndexFields>")
            If (_IsCustomerNumber) Then
                s1.Write(vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<IndexField Name=""Customer Account"" Value = """ & sAccountNumber & """ />")
            Else
                s1.Write(vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<IndexField Name=""Bill Account"" Value = """ & sAccountNumber & """ />")
            End If
            s1.Write(vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<IndexField Name=""Correspondence Type"" Value = """ & corrType & """ />")
            s1.Write(vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<IndexField Name=""Correspondence Source"" Value = """ & corrSource & """ />")
            s1.Write(vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<IndexField Name=""Comments"" Value = """ & strComments & """ />")
            s1.Write(vbTab & vbTab & vbTab & vbTab & vbTab & "</IndexFields>")
            s1.Write(vbTab & vbTab & vbTab & vbTab & vbTab & "<Pages>")
            s1.Write(vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<Page ImportFileName=""" & value & """/>")
            s1.Write(vbTab & vbTab & vbTab & vbTab & vbTab & "</Pages>")
            s1.Write(vbTab & vbTab & vbTab & vbTab & "</Document>")
        Next

        'End XML File
        s1.Write(vbTab & vbTab & vbTab & "</Documents>")
        s1.Write(vbTab & vbTab & "</Batch>")
        s1.Write(vbTab & "</Batches>")
        s1.Write("</ImportSession>")
        s1.Close()
        fs1.Close()
    End Sub

End Module
