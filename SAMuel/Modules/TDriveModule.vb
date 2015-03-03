Imports System.Data.OleDb
Imports System.IO
Imports Microsoft.Office.Interop.Word
Imports System.Threading
Imports Microsoft.Office.Interop.Outlook
Imports System.Text.RegularExpressions

Namespace Modules
    Module TDriveModule
        Private _con As OleDbConnection = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + My.Settings.DatabaseFile)
        Public Enum DeliveryType As Byte
            Err = 0
            Email = 1
            Fax = 2
            Mail = 3
        End Enum

        Public Enum DPAType As Byte
            Err = 0
            Active = 1
            Cutin = 2
            AcctInit = 3
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
            Public TimeToSend As TimeSpan

            'Internal Data
            Public Skip As Boolean
            Public Sent As Boolean

            Public Sub New(ByVal sFile As String)
                Me.SourceFile = sFile
                Me.CreationTime = RemoveMilliseconds(File.GetCreationTime(sFile))
                Me.AccountNumber = RegexAcc(Path.GetFileName(SourceFile), "\d{5}-\d{5}")
                Me.Skip = False
                Me.Sent = False
            End Sub


            ''' <summary>
            '''     Parses the word document for the Send To info.
            ''' </summary>
            ''' <param name="wordApplication">Reference to an open word application object.</param>
            ''' <remarks></remarks>
            Public Sub ExtractDetailsFromDoc(ByRef wordApplication As Microsoft.Office.Interop.Word.Application)
                'Open the word document associated with the DPA Struct.
                wordApplication.Visible = False
                Dim objWdDoc As New Document
                objWdDoc = wordApplication.Documents.Open(FileName:=Me.SourceFile, ConfirmConversions:=False, ReadOnly:=True)
                With objWdDoc
                    'Skip the document if it doesn't fit the design constraints of a standard dpa form.
                    If objWdDoc.Tables.Count <> 1 Then
                        Me.Skip = True
                        LogAction(0,
                                  Me.SourceFile & " was skipped. Invalid table count:" &
                                  objWdDoc.Tables.Count.ToString())
                        objWdDoc.Close()
                        Exit Sub
                    End If

                    'Extract data from the DPA header table
                    With objWdDoc.Tables(1)
                        'Store the sendTo table cell to a temp string
                        Dim tempSendToCell As String = .Cell(1, 1).Range.Text

                        'Determine the delivery method
                        Dim deliveryTuple As Tuple(Of DeliveryType, String) = DetermineDeliveryMethod(tempSendToCell)
                        If IsNothing(deliveryTuple) Then
                            Me.Skip = True
                        Else
                            Me.DeliveryMethod = deliveryTuple.Item1
                            Me.SendTo = DataScrubber(deliveryTuple.Item2)
                        End If

                        For Each c As Char In Me.SendTo
                            Debug.Print(c + " " + Asc(c).ToString)
                        Next

                        'Cutin/Active/AcctInit
                        Dim sType As String = .Cell(1, 2).Range.Text
                        If InStr(sType, "Active", CompareMethod.Text) Then
                            Me.Type = DPAType.Active
                        ElseIf InStr(sType, "Cut-In", CompareMethod.Text) Then
                            Me.Type = DPAType.Cutin
                        ElseIf InStr(sType, "Account Initiation", CompareMethod.Text) Then
                            Me.Type = DPAType.AcctInit
                        Else
                            'Error
                            Me.Type = DPAType.Err
                            Me.Skip = True
                        End If

                        'Customer Name with string cleanse and formating
                        Dim tempStr = CleanInput(.Cell(2, 1).Range.Text) _
                        'Extract customer name from the table and cleanse the string.
                        tempStr = Replace(tempStr, "Customer Name", String.Empty) 'Remove 'Customer Name' wording
                        Me.CustomerName = DataScrubber(StrConv(tempStr, VbStrConv.Uppercase))

                        Me.AccountNumber = RegexAcc(.Cell(2, 2).Range.Text, "\d{5}-\d{5}")  'Account Number

                        ''Commented out for future reference or use.
                        'address = .Cell(3, 1).Range.Text 'Service Address
                        'dateOffer = .Cell(3, 2).Range.Text  'Date Offered
                        ''
                    End With

                    'Print to PDF or Printer depending on type

                    'Get the desired printer and set it as active printer in Word.
                    Dim desiredPrinter As String
                    If Me.DeliveryMethod = DeliveryType.Email Then
                        desiredPrinter = My.Settings.PDF_PrinterName
                        Me.FileToSend = TDrive_FOLDER & Path.GetFileNameWithoutExtension(Me.SourceFile) & ".pdf"
                    ElseIf Me.DeliveryMethod = DeliveryType.Mail Then
                        desiredPrinter = My.Settings.Physical_PrinterName
                    Else
                        desiredPrinter = My.Settings.PDF_PrinterName
                        Me.Skip = True
                    End If

                    If (Me.Skip = False) Then
                        Try
                            wordApplication.ActivePrinter = desiredPrinter
                        Catch ex As System.Exception
                            MsgBox("Printer error. Is the " + desiredPrinter + " printer installed?", MsgBoxStyle.Critical, "Printer error!")
                            Me.Skip = True
                            objWdDoc.Close()
                            Return
                        End Try


                        objWdDoc.PrintOut()
                        Thread.Sleep(3000)
                    End If

                    'If Not File.Exists(Me.FileToSend) Then
                    '    MsgBox("Expected PDF was not created. Check the settings folder for SAMuel and PDF995.")
                    '    Me.Skip = True
                    'End If
                    'Release document
                    objWdDoc.Close()
                End With
            End Sub


            ''' <summary>
            '''     Marks the DPA as sent out and calls cleanup routines.
            ''' </summary>
            ''' <remarks></remarks>
            Sub Complete()
                Me.CompletionTime = RemoveMilliseconds(Now())
                Me.TimeToSend = Me.CompletionTime - Me.CreationTime
            End Sub
        End Structure

        Public Sub ProcessFiles(sFiles() As String)
            Dim objWord As Microsoft.Office.Interop.Word.Application
            Dim olApp As Microsoft.Office.Interop.Outlook.Application =
                    New Microsoft.Office.Interop.Outlook.Application
            Dim dpaList As New List(Of DPA)

            'Initiate word application object and minimize it
            objWord = CreateObject("Word.Application")
            objWord.WindowState = WdWindowState.wdWindowStateMinimize
            'Set active printer to PDF
            'objWord.ActivePrinter = "PDF995"

            'Setup the progress bar.
            FrmMain.ProgressBar.Maximum = sFiles.Length + 1
            FrmMain.ProgressBar.Value = 0

            'For each DPA created from each file within the list of files...
            For Each newDPA As DPA In From sFile As String In sFiles Select New DPA(sFile)
                Try
                    'Skip Hidden Files.
                    If ((File.GetAttributes(newDPA.SourceFile) & FileAttributes.Hidden) = FileAttributes.Hidden) Then
                        newDPA.Skip = True
                        Continue For
                    End If

                    FrmMain.lblStatus.Text = "Parsing " + Path.GetFileName(newDPA.SourceFile)
                    FrmMain.Refresh()

                    newDPA.ExtractDetailsFromDoc(objWord)

                    FrmMain.ProgressBar.Value += 1


                    'Add each DPA to the list view for visual verification.
                    Dim lvi As ListViewItem = New ListViewItem(StrConv(newDPA.Type.ToString, VbStrConv.ProperCase))
                    lvi.SubItems.Add(newDPA.SendTo)
                    lvi.SubItems.Add(newDPA.AccountNumber)
                    lvi.SubItems.Add(newDPA.CustomerName)

                    'If it's marked as skip, Add to list as Red and Move to next file i
                    If newDPA.Skip Then
                        newDPA.Sent = False
                        lvi.Tag = newDPA
                        lvi.ForeColor = Color.Red
                        FrmMain.lvTDriveFiles.Items.Add(lvi)
                        FrmMain.Refresh()
                        Continue For
                    End If

                    If newDPA.DeliveryMethod = DeliveryType.Mail Then
                        newDPA.Sent = True
                        lvi.ForeColor = SystemColors.Highlight
                        newDPA.Complete()
                    Else
                        'Email the DPA and mark as complete if successful
                        FrmMain.lblStatus.Text = "Emailing the DPA..."
                        FrmMain.Refresh()

                        If SendEmail(newDPA, olApp) Then
                            newDPA.Sent = True
                            lvi.ForeColor = SystemColors.ControlText 
                            newDPA.Complete()
                        Else
                            ''TODO Log failed send email
                            newDPA.Sent = False
                            lvi.ForeColor = Color.Red
                            LogAction(0, newDPA.SourceFile + " | " + newDPA.FileToSend)
                        End If


                    End If

                    'Add item to listview
                    lvi.Tag = newDPA
                    dpaList.Add(newDPA)
                    FrmMain.lvTDriveFiles.Items.Add(lvi)
                    FrmMain.Refresh()

                    ''Move the Source DPA file if it was sent.
                    If newDPA.Sent Then
                        AddRecordToDatabase(newDPA)
                        If newDPA.Type = DPAType.Active Then
                            MoveEmailedFile(newDPA.SourceFile, My.Settings.EmailedActiveMoveFolder)
                        ElseIf newDPA.Type = DPAType.Cutin Then
                            MoveEmailedFile(newDPA.SourceFile, My.Settings.EmailedCutinMoveFolder)
                        ElseIf newDPA.Type = DPAType.AcctInit Then
                            MoveEmailedFile(newDPA.SourceFile, My.Settings.EmailedAccInitMoveFolder)
                        Else
                            MoveEmailedFile(newDPA.SourceFile, My.Settings.EmailedActiveMoveFolder)
                        End If
                    End If

                Catch ex As System.Exception
                    Continue For
                End Try
            Next

            'UI Update
            FrmMain.lblStatus.Text = "DONE!"
            FrmMain.ProgressBar.Value += 1
            'Cleanup
            objWord.Quit()
            objWord = Nothing
            dpaList.Clear()
        End Sub

        Private Sub AddRecordToDatabase(ByVal dpaDoc As DPA)
            Const strDPACommand As String =
            "INSERT INTO DeferredPaymentAgreements (DPAType,DeliveryMethod,AccountNumber,SentTo,CustomerName,SentTime,FileCreated,File) " +
            "VALUES (@dPAType,@delivery,@account,@sendTo,@customerName,@timeSent,@fileCreationTime,@document)"
            _con.Open()
            Dim cmdInsert As OleDbCommand = New OleDbCommand(strDPACommand, _con)

            cmdInsert.Parameters.AddWithValue("@dPAType", dpaDoc.Type.ToString())
            cmdInsert.Parameters.AddWithValue("@delivery", dpaDoc.DeliveryMethod.ToString())
            cmdInsert.Parameters.AddWithValue("@account", dpaDoc.AccountNumber)
            cmdInsert.Parameters.AddWithValue("@sendTo", dpaDoc.SendTo)
            cmdInsert.Parameters.AddWithValue("@customerName", dpaDoc.CustomerName)
            cmdInsert.Parameters.AddWithValue("@timeSent", dpaDoc.CompletionTime)
            cmdInsert.Parameters.AddWithValue("@fileCreationTime", dpaDoc.CreationTime)
            cmdInsert.Parameters.AddWithValue("@document", dpaDoc.SourceFile)
            cmdInsert.ExecuteNonQuery()

            _con.Close()

        End Sub

        Private Function SendEmail(ByRef outDPA As DPA, ByRef olApp As Microsoft.Office.Interop.Outlook.Application) _
            As Boolean
            Try
                Dim olEmail As MailItem = olApp.CreateItem(OlItemType.olMailItem)
                Dim recipents As Recipients = olEmail.Recipients
                recipents.Add(outDPA.SendTo)

                'Determine the sending address.
                If outDPA.Type = DPAType.Active Then
                    olEmail.SentOnBehalfOfName = My.Settings.ACTIVE_EMAIL
                ElseIf outDPA.Type = DPAType.Cutin Then
                    olEmail.SentOnBehalfOfName = My.Settings.CUTIN_EMAIL
                ElseIf outDPA.Type = DPAType.AcctInit Then
                    olEmail.SentOnBehalfOfName = My.Settings.ACCINIT_EMAIL
                Else
                    Return False
                End If

                'Create the email and save as draft
                olEmail.Subject = outDPA.AccountNumber & " Deferred Payment Agreement"
                olEmail.Body = My.Settings.Email_Body
                olEmail.Body += vbCrLf + vbCrLf + vbCrLf
                olEmail.BodyFormat = OlBodyFormat.olFormatRichText
                olEmail.Attachments.Add(outDPA.FileToSend)
                olEmail.Save()
                ''TODO Change to .Send once security issue is addressed.
                Thread.Sleep(1000)
                ''If olEmail.Sent Then
                Return True
                ''Else
                ''Return False
                ''End If

            Catch ex As System.Exception
                LogAction(0, ex.Message)
                Return False
            End Try
        End Function

        Private Sub MoveEmailedFile(sourceFile As String, moveToFolder As String)
            Try

                Dim destination As String = moveToFolder + Path.GetFileName(sourceFile)
                If File.Exists(destination) Then
                    File.Delete(destination)
                End If
                File.Move(sourceFile, destination)

            Catch

            End Try
        End Sub

        Private Function DetermineDeliveryMethod(sendToField As String) As Tuple(Of DeliveryType, String)
            'Check for email address.
            Const rgxEmailPattern As String =
                "[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?"
            Dim rgx As Regex = New Regex(rgxEmailPattern, RegexOptions.IgnoreCase)
            Dim emailMatch As Match = rgx.Match(sendToField)
            If (emailMatch.Success) Then
                Return Tuple.Create(DeliveryType.Email, emailMatch.Value.ToLower()) 'Return email address lowercase
            End If

            'Check for a mailing address.
            'Look for space, two letters, space then 5 digits
            'ex. ' NY 13219'
            Const rgxAddressPattern As String = "\s[a-z]{2}\s\d{5}"
            rgx = New Regex(rgxAddressPattern, RegexOptions.IgnoreCase)
            Dim mailMatch As Match = rgx.Match(sendToField)
            If (mailMatch.Success) Then
                Return Tuple.Create(DeliveryType.Mail, RemoveSendToText(sendToField).ToUpper)
            End If
            Return Tuple.Create(DeliveryType.Err, RemoveSendToText(sendToField).ToUpper)
        End Function

        'Cleans the field of excel/table characters
        Private Function DataScrubber(strData As String) As String
            'Removes carriage return, Line Feed and Bell.
            Return strData.Replace(vbLf, String.Empty).Replace(vbCr, String.Empty).Replace(Chr(7), String.Empty).Trim()
        End Function

        'Removes each possible option when mailing.
        'Can't trust user selection.
        Private Function RemoveSendToText(strField As String) As String
            Dim junkToRemove As String() = {"Email to:", "Mail to:", "Fax to:"}
            For Each s As String In junkToRemove
                strField = strField.Replace(s, String.Empty)
            Next
            Return (strField.Trim())
        End Function

    End Module
End Namespace