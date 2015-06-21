Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports System.Text
Imports System.Threading.Tasks

Namespace Modules

    Structure ContactStruct
        Dim DeliveryMethod As String
        Dim SentTo As String
        Dim DPAType As String
    End Structure

    Module AddContact
        Private _con As OleDbConnection
        Friend contactDataSet As DataSet

        Public UserCanceledContactAdding As Boolean

        Public Function RunScript(strBillAccount As String, strContact As String, strContactType As String) As Boolean
            Dim arguments As String = String.Format(" /account {0} /contact ""{1}"" /type ""{2}""", strBillAccount, strContact, strContactType)
            Dim strProgram As String = Application.StartupPath + "\addContact.exe"

            'Call the script and wait.
            Dim p As New Process
            Dim psi As New ProcessStartInfo(strProgram, arguments)
            p.StartInfo = psi
            p.Start()
            p.WaitForExit()
            If (p.ExitCode = 100) Then
                'Success
                Return True
            Else
                Return False
            End If
        End Function

        Public Function GetAccountsNeedingContacts(databaseFile As String) As Boolean
            'Check if the connection/database exists
            If (Not File.Exists(databaseFile)) Then
                MsgBox(String.Format("Database file, {0}, was Not found. Please check your connection And/Or path And Try again.", databaseFile), MsgBoxStyle.Critical, "Database Not found!")
                _con = Nothing
                Return False
            End If

            If (IsNothing(_con)) Then
                _con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + databaseFile)
            End If
            Const strCmd As String = "Select DeferredPaymentAgreements.AccountNumber, DeferredPaymentAgreements.DPAType, DeferredPaymentAgreements.SentTo, DeferredPaymentAgreements.DeliveryMethod, DeferredPaymentAgreements.ContactAdded, DeferredPaymentAgreements.Key  FROM DeferredPaymentAgreements WHERE (((DeferredPaymentAgreements.ContactAdded)=False));"



            'Connect to database, run query and store the dataset
            _con.Open()  'open up a connection to the database
            Dim cmd As OleDbCommand = New OleDbCommand(strCmd, _con)
            contactDataSet = New DataSet
            Dim da As OleDbDataAdapter = New OleDbDataAdapter(cmd.CommandText, _con)
            Dim row As System.Data.DataRow
            da.Fill(contactDataSet, "ContactsNeeded") 'Fill the dataset, contactDataSet, with the above SELECT statement
            _con.Close()

            'Fill the DataGridView
            'NOTE Verify select command order matches datagrid column order.
            For Each row In contactDataSet.Tables("ContactsNeeded").Rows
                ' FrmMain.DataGridView.Rows.Add(row.ItemArray)
            Next

            Return True
        End Function

        Public Sub AddContactsToDatasetAccounts()
			Const contactType = "Miscellaneous Collections"	'Always this type (For now)
			Dim ContactsCompletedSinceUpdate As Integer = 0
			Dim bFirstSuccess As Boolean = True
            If (IsNothing(contactDataSet)) Then
                Return
            End If
            If (contactDataSet.Tables.Count = 1) Then
                'Start the 
                Dim updateString As StringBuilder = New StringBuilder
                updateString.Append("UPDATE DeferredPaymentAgreements Set DeferredPaymentAgreements.ContactAdded = True WHERE")

                'Add the contact for each account in the dataset.
                For Each row As DataRow In contactDataSet.Tables("ContactsNeeded").Rows

                    'Skip the row if it was deleted by the user.
                    If row.RowState = DataRowState.Deleted Then Continue For

                    'Collect Account Number, DPA Type, Method of delivery and where it was sent.
                    Dim rsKey As Integer = row("Key")
                    Dim accountNumber As String = row("AccountNumber")

                    'Generate the contact string.
                    Dim contactInfo As ContactStruct = New ContactStruct()
                    contactInfo.DeliveryMethod = row("DeliveryMethod")
                    contactInfo.DPAType = row("DPAType")
                    contactInfo.SentTo = row("SentTo")
                    Dim contactString As String = GenerateContactString(contactInfo)

                    'Run the script to add the contact
                    If (RunScript(accountNumber, contactString, contactType)) Then
						ContactsCompletedSinceUpdate += 1
                        'Update the item in the dataset.
                        row("ContactAdded") = True
						If (bFirstSuccess) Then
							updateString.Append(" [DeferredPaymentAgreements].[Key]=" & rsKey.ToString)
							bFirstSuccess = False 'No longer the first success
                        Else
							updateString.Append(" Or [DeferredPaymentAgreements].[Key]=" & rsKey.ToString)
						End If

						'Update database after 25 records and reset the string
						''Expression too complex when over 100
						If (ContactsCompletedSinceUpdate >= 25) Then
							updateString.Append(";")
							UpdateDatabase(updateString.ToString)
							updateString.Clear()
							updateString.Append("UPDATE DeferredPaymentAgreements Set DeferredPaymentAgreements.ContactAdded = True WHERE")
							bFirstSuccess = True
							ContactsCompletedSinceUpdate = 0
						End If
					End If
				Next

                'Finish Update Query String
                If (bFirstSuccess = False) Then
                    updateString.Append(";")
					UpdateDatabase(updateString.ToString)
				Else
                    'No contacts were added successfully
                    'TODO Give user feedback?
                    Return
                End If
            End If
        End Sub

        Private Sub UpdateDatabase(ByVal updateString As String)
            If (IsNothing(_con)) Then
                _con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + My.Settings.DatabaseFile)
            End If

            'Connect to database, run query and store the dataset
            _con.Open()  'open up a connection to the database
            Dim cmd As OleDbCommand = New OleDbCommand(updateString, _con)
            cmd.ExecuteNonQuery()
            _con.Close()
        End Sub

        Private Function GenerateContactString(contactInfo As ContactStruct) As String
            Dim contactString As String

            'Create the contact string.
            ''Example: Emailed Active DPA to test@test.com
            contactString = String.Format("{0}ed {1} DPA To {2}", contactInfo.DeliveryMethod, contactInfo.DPAType, contactInfo.SentTo)
            Return contactString
        End Function

        Public Sub DeleteContact(key As Integer)
            Dim query As String =
            "UPDATE DeferredPaymentAgreements Set DeferredPaymentAgreements.ContactAdded = True WHERE [DeferredPaymentAgreements].[Key]=" & key.ToString
            UpdateDatabase(query)
        End Sub

    End Module
End Namespace