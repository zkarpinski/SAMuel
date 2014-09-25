Imports System.IO

Namespace Modules

    Module KofaxModule

        ''' <summary>
        ''' Creates an XML file which is read by Kofax Import Connector to create a batch
        ''' </summary>
        ''' <param name="docList">
        ''' List of per document tiff files stored in another list
        ''' </param>
        Sub CreateXML(ByRef docList As List(Of List(Of String)), batchName As String)
            '' Deprecated overload. Removing soon.
            Dim sPath As String
            Dim sFile As String
            Dim sList As List(Of String)
            Dim batchClass As String

            sPath = "C:\ACXMLAID\" 'Default Kofax Import Connector watch folder
            sFile = sPath & batchName & "-import.xml"
            batchClass = "eCorrespondence Fxd Pg"

            'Creates the directory if it doesn't exist
            If Not Directory.Exists(sPath) Then
                Directory.CreateDirectory(sPath)
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
            Dim isCustomerNumber As Boolean
            Dim batchClass As String
            Dim accountDictionary As Dictionary(Of String, List(Of String)) = New Dictionary(Of String, List(Of String))


            'Write each tiff file as a document
            For Each value In docList

                sAccountNumber = RegexAcc(value, "\d{5}-\d{5}")
                If sAccountNumber = "X" Then
                    sAccountNumber = RegexAcc(value, "\d{10}")
                    If sAccountNumber = "X" Then
                        sAccountNumber = RegexAcc(value, "\d{9}")
                        If sAccountNumber = "X" Then
                            sAccountNumber = "UNKNOWN"
                        End If
                    End If
                End If

                'Remove Hyphen from account number
                sAccountNumber = sAccountNumber.Replace("-", "")

                If accountDictionary.ContainsKey(sAccountNumber) Then
                    'Add the file to the existing account number list
                    accountDictionary.Item(sAccountNumber).Add(value)
                Else
                    'Add the file to a list and add the account with the list to the dictionary.
                    Dim fileList As List(Of String) = New List(Of String)
                    fileList.Add(value)
                    accountDictionary.Add(sAccountNumber, fileList)
                End If

            Next

            sPath = "C:\ACXMLAID\" 'Default Kofax Import Connector watch folder
            sFile = sPath & batchName & "-import.xml"
            batchClass = "eCorrespondence Fax"

            'Creates the directory if it doesn't exist
            If Not Directory.Exists(sPath) Then
                Directory.CreateDirectory(sPath)
            End If

            'Creates or Overwrites the XML file
            Dim fs1 As FileStream = New FileStream(sFile, FileMode.Create, FileAccess.Write)
            Dim s1 As StreamWriter = New StreamWriter(fs1)

            'XML Header
            s1.Write("<ImportSession>")
            s1.Write(vbTab & "<Batches>")
            s1.Write(vbTab & vbTab & "<Batch BatchClassName=""" & batchClass & """ EnableSingleDocProcessing = ""0"" Name=""" & batchName & """>")
            s1.Write(vbTab & vbTab & vbTab & "<Documents>")

            'Iterate through the dictonary
            Dim acc As KeyValuePair(Of String, List(Of String))
            For Each acc In accountDictionary
                Dim fileEntry As String
                'Get the account number
                Dim account As String = acc.Key

                'Determine the account type. Customer/Bill Account
                If account.Length = 10 Then
                    'Is a bill account number
                    isCustomerNumber = False
                ElseIf account.Length = 9 Then
                    'Is  a customer number
                    isCustomerNumber = True
                ElseIf account = "UNKNOWN" Then
                    'There was no account number
                    'Doesn't matter what it's inserted into.
                    isCustomerNumber = False
                Else
                    'This case shouldn't be reached but..
                    isCustomerNumber = False
                    LogAction(99, String.Format("Kofax it account number type else case reached. Account: {0} File: {1}", account, acc.Value.Item(0)))
                End If

                'Write the XML File
                'TODO Use DOM
                s1.Write(vbTab & vbTab & vbTab & vbTab & "<Document FormType = ""eCorrespondence Fax"">")
                s1.Write(vbTab & vbTab & vbTab & vbTab & vbTab & "<IndexFields>")
                If (isCustomerNumber) Then
                    s1.Write(vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<IndexField Name=""Customer Account"" Value = """ & account & """ />")
                Else
                    s1.Write(vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<IndexField Name=""Bill Account"" Value = """ & account & """ />")
                End If
                s1.Write(vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<IndexField Name=""Correspondence Type"" Value = """ & corrType & """ />")
                s1.Write(vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<IndexField Name=""Correspondence Source"" Value = """ & corrSource & """ />")
                s1.Write(vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<IndexField Name=""Comments"" Value = """ & strComments & """ />")
                s1.Write(vbTab & vbTab & vbTab & vbTab & vbTab & "</IndexFields>")
                s1.Write(vbTab & vbTab & vbTab & vbTab & vbTab & "<Pages>")

                'Add each file associated to the account, as a page in kofax. 
                ''*NOTE* A page in kofax can be one or more actual pages.
                For Each fileEntry In acc.Value
                    s1.Write(vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "<Page ImportFileName=""" & fileEntry & """/>")
                Next

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
End Namespace