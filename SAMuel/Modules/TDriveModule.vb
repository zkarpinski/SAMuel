Imports Microsoft.Office.Interop
Imports System.IO
Imports System.Linq

Module TDriveModule

    Public Enum DeliveryType As Byte
        ERR = 0
        Email = 1
        Fax = 2
        Mail = 3
    End Enum

    Public Enum DPAType As Byte
        ERR = 0
        ACTIVE = 1
        CUTIN = 2
        ACCINIT = 3
    End Enum

    Public Structure DPA
        ' Required base information.
        Public SourceFile As String
        Public Type As DPAType
        Public AccountNumber As String
        Public CustomerName As String
        Public FileToSend As String
        Public DeliveryMethod As DeliveryType
        Public SendTo As String 'Can be Fax Number, Email, or Address.

        'Metrics Data
        Public CreationTime As Date
        Public CompletionTime As Date
        Public TimeToSend As Date

        'Internal Data
        Public SKIP As Boolean

        Public Sub New(ByVal sFile As String)
            Me.SourceFile = sFile
            Me.CreationTime = File.GetCreationTime(sFile)
            Me.AccountNumber = RegexAcc(Path.GetFileName(SourceFile), "\d{5}-\d{5}")
            Me.SKIP = False
        End Sub

        ''' <summary>
        ''' Parses the word document for the Send To info.
        ''' </summary>
        ''' <param name="wordApplication">Reference to an open word application object.</param>
        ''' <remarks></remarks>
        Public Sub ExtractDetailsFromDoc(ByRef wordApplication As Word.Application)
            'Open the word document associated with the DPA Struct.
            wordApplication.Visible = False
            Dim objWdDoc As New Word.Document
            objWdDoc = wordApplication.Documents.Open(FileName:=Me.SourceFile, ConfirmConversions:=False)
            With objWdDoc
                'Skip the document if it doesn't fit the design constraints of a standard dpa form.
                If objWdDoc.Tables.Count <> 1 Then
                    objWdDoc.Close()
                    Me.SKIP = True
                    LogAction(0, Me.SourceFile & " was skipped. Invalid table count:" & objWdDoc.Tables.Count.ToString())
                    Exit Sub
                End If

                'Extract data from the DPA header table
                With objWdDoc.Tables(1)
                    'Email Address
                    Me.SendTo = RegexAcc(.Cell(1, 1).Range.Text, _
                                         "[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?")

                    'TODO: Add handling for address to mail out

                    'Cutin/Active
                    If InStr(.Cell(1, 2).Range.Text, "Active", CompareMethod.Text) Then
                        Me.Type = DPAType.ACTIVE
                    ElseIf InStr(.Cell(1, 2).Range.Text, "Cutin", CompareMethod.Text) Then
                        Me.Type = DPAType.CUTIN
                    Else
                        'Error
                        Me.Type = DPAType.ERR
                        Me.SKIP = True
                    End If

                    'Customer Name with string cleanse and formating
                    Dim tempStr = GlobalModule.CleanInput(.Cell(2, 1).Range.Text) 'Extract customer name from the table and cleanse the string.
                    tempStr = Replace(tempStr, "Customer Name", "") 'Remove 'Customer Name' wording
                    Me.CustomerName = StrConv(tempStr, VbStrConv.ProperCase)  'Capitalize the first letters

                    Me.AccountNumber = RegexAcc(.Cell(2, 2).Range.Text, "\d{5}-\d{5}")  'Account Number

                    ''Commented out for future reference or use.
                    'address = .Cell(3, 1).Range.Text 'Service Address
                    'dateOffer = .Cell(3, 2).Range.Text  'Date Offered
                    ''
                End With

                'TODO: Print to PDF or Printer depending on type
                ' objWdDoc.PrintOut(PrintToFile:=True, OutputFileName:=sDestination & sFileName & ".pdf")

                'Release document
                objWdDoc.Close()
            End With
        End Sub

    End Structure


    Public Sub ProcessFiles(sFiles() As String)
        Dim objWord As Word.Application
        Dim DPA_List As New List(Of DPA)

        'Initiate word application object and minimize it
        objWord = CreateObject("Word.Application")
        objWord.WindowState = Word.WdWindowState.wdWindowStateMinimize
        'Set active printer to PDF
        'objWord.ActivePrinter = "PDF995"

        'Setup the progress bar.
        frmMain.ProgressBar.Maximum = sFiles.Length + 1
        frmMain.ProgressBar.Value = 0
        frmMain.lblStatus.Text = "Parsing DPA files..."

        'Create list of DPAs
        For Each value In sFiles
            Dim newDPA As DPA = New DPA(value)
            newDPA.ExtractDetailsFromDoc(objWord)
            DPA_List.Add(newDPA)
            frmMain.ProgressBar.Value += 1
            frmMain.Refresh()
        Next

        'UI Update
        frmMain.lblStatus.Text = "Updating UI..."
        frmMain.Refresh()

        For Each letter As DPA In DPA_List
            'Add each DPA to the list view for visual verification.
            Dim lvi As ListViewItem = New ListViewItem(StrConv(letter.Type.ToString, VbStrConv.ProperCase))
            lvi.SubItems.Add(letter.SendTo)
            lvi.SubItems.Add(letter.AccountNumber)
            lvi.SubItems.Add(letter.CustomerName)
            lvi.Tag = letter.SourceFile 'Allows the file to be opened when clicked.
            frmMain.lvTDriveFiles.Items.Add(lvi)
        Next

        'UI Update
        frmMain.lblStatus.Text = "DONE!"
        frmMain.ProgressBar.Value += 1
        'Cleanup
        objWord.Quit()
        objWord = Nothing
    End Sub
End Module
