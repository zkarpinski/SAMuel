Imports System.IO
Imports Microsoft.Office.Interop.Word
Imports System.Threading
Imports Microsoft.Office.Interop.Outlook

Namespace Modules
    Module TDriveModule
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
            Accinit = 3
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
                Me.CreationTime = File.GetCreationTime(sFile)
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
                objWdDoc = wordApplication.Documents.Open(FileName:=Me.SourceFile, ConfirmConversions:=False)
                With objWdDoc
                    'Skip the document if it doesn't fit the design constraints of a standard dpa form.
                    If objWdDoc.Tables.Count <> 1 Then
                        Me.SKIP = True
                        LogAction(0,
                                  Me.SourceFile & " was skipped. Invalid table count:" &
                                  objWdDoc.Tables.Count.ToString())
                        objWdDoc.Close()
                        Exit Sub
                    End If

                    'Extract data from the DPA header table
                    With objWdDoc.Tables(1)
                        'Store the table cell to a temp string as lowercase.
                        Dim tempSendToCell As String = .Cell(1, 1).Range.Text
                        tempSendToCell = tempSendToCell.ToLower
                        'Email Address
                        Me.SendTo = RegexAcc(tempSendToCell,
                                             "[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?")

                        ' TODO Add handling for mailing address and fax number.

                        'Determine Delivery Method
                        If Me.SendTo.Contains("@") Then
                            Me.DeliveryMethod = DeliveryType.Email
                        Else
                            Me.Skip = True
                        End If

                        'Cutin/Active
                        Dim sType As String = .Cell(1, 2).Range.Text
                        If InStr(sType, "Active", CompareMethod.Text) Then
                            Me.Type = DPAType.Active
                        ElseIf InStr(sType, "Cut-In", CompareMethod.Text) Then
                            Me.Type = DPAType.Cutin
                        ElseIf InStr(sType, "Account Initiation", CompareMethod.Text) Then
                            Me.Type = DPAType.Accinit
                        Else
                            'Error
                            Me.Type = DPAType.Err
                            Me.Skip = True
                        End If

                        'Customer Name with string cleanse and formating
                        Dim tempStr = CleanInput(.Cell(2, 1).Range.Text) _
                        'Extract customer name from the table and cleanse the string.
                        tempStr = Replace(tempStr, "Customer Name", "") 'Remove 'Customer Name' wording
                        Me.CustomerName = StrConv(tempStr, VbStrConv.ProperCase)  'Capitalize the first letters

                        Me.AccountNumber = RegexAcc(.Cell(2, 2).Range.Text, "\d{5}-\d{5}")  'Account Number

                        ''Commented out for future reference or use.
                        'address = .Cell(3, 1).Range.Text 'Service Address
                        'dateOffer = .Cell(3, 2).Range.Text  'Date Offered
                        ''
                    End With

                    'Print to PDF or Printer depending on type
#If CONFIG = "Release" Then
            'Set active printer to PDF995
            Try
                wordApplication.ActivePrinter = "PDF995"
                    Catch
                        MsgBox("Printer error. Is the 'PDF995' printer installed?", MsgBoxStyle.Critical)
                        Me.Skip = True
                        objWdDoc.Close()
                        Return
            End Try
#End If
                    objWdDoc.PrintOut()
                    Thread.Sleep(3000)
                    Me.FileToSend = TDrive_FOLDER & Path.GetFileNameWithoutExtension(Me.SourceFile) & ".pdf"
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
                Me.CompletionTime = Now()
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

                    'Move to next file if it's marked as skip
                    If newDPA.Skip Then Continue For

                    'Add each DPA to the list view for visual verification.
                    Dim lvi As ListViewItem = New ListViewItem(StrConv(newDPA.Type.ToString, VbStrConv.ProperCase))
                    lvi.SubItems.Add(newDPA.SendTo)
                    lvi.SubItems.Add(newDPA.AccountNumber)
                    lvi.SubItems.Add(newDPA.CustomerName)
                    lvi.Tag = newDPA

                    'Email the DPA and move if sucessful.
                    FrmMain.lblStatus.Text = "Emailing the DPA..."
                    FrmMain.Refresh()

                    If SendEmail(newDPA, olApp) Then
                        newDPA.Sent = True
                        lvi.ForeColor = Color.Black
                        newDPA.Complete()
                    Else
                        ''TODO Log failed send email
                        newDPA.Sent = False
                        lvi.ForeColor = Color.Red
                        LogAction(0, newDPA.SourceFile + " | " + newDPA.FileToSend)
                    End If

                    lvi.Tag = newDPA
                    dpaList.Add(newDPA)
                    FrmMain.lvTDriveFiles.Items.Add(lvi)
                    FrmMain.Refresh()

                    ''Move the Source DPA file if it was sent.
                    If newDPA.Sent Then
                        If newDPA.Type = DPAType.Active Then
                            MoveEmailedFile(newDPA.SourceFile, My.Settings.EmailedActiveMoveFolder)
                        ElseIf newDPA.Type = DPAType.Cutin Then
                            MoveEmailedFile(newDPA.SourceFile, My.Settings.EmailedCutinMoveFolder)
                        ElseIf newDPA.Type = DPAType.Accinit Then
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
            frmMain.lblStatus.Text = "DONE!"
            frmMain.ProgressBar.Value += 1
            'Cleanup
            objWord.Quit()
            objWord = Nothing
            dpaList.Clear()
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
                ElseIf outDPA.Type = DPAType.Accinit Then
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
    End Module
End Namespace