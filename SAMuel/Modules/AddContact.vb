Imports System.Data
Imports System.Data.OleDb
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
        Private _ds As DataSet

        Public UserCanceledContactAdding As Boolean

        Public Function RunScript(strBillAccount As String, strContact As String) As Boolean
            Dim arguments As String = String.Format(" /account {0} /contact ""{1}", strBillAccount, strContact)

            'Call the script and wait.
            Dim p As New Process
            Dim psi As New ProcessStartInfo(Application.StartupPath + "\addContact.exe", arguments)
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

        Public Sub GetAccountsNeedingContacts(databaseFile As String)
            If (IsNothing(_con)) Then
                _con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + databaseFile)
            End If
            Const strCmd As String = "SELECT DeferredPaymentAgreements.AccountNumber, DeferredPaymentAgreements.DPAType, DeferredPaymentAgreements.SentTo, DeferredPaymentAgreements.DeliveryMethod, DeferredPaymentAgreements.ContactAdded, DeferredPaymentAgreements.Key  FROM DeferredPaymentAgreements WHERE (((DeferredPaymentAgreements.ContactAdded)=False));"


            'Connect to database, run query and store the dataset
            _con.Open()  'open up a connection to the database
            Dim cmd As OleDbCommand = New OleDbCommand(strCmd, _con)
            _ds = New DataSet
            Dim da As OleDbDataAdapter = New OleDbDataAdapter(cmd.CommandText, _con)
            Dim row As System.Data.DataRow
            da.Fill(_ds, "ContactsNeeded") 'Fill the dataset, _ds, with the above SELECT statement
            _con.Close()

            'Fill the DataGridView
            'NOTE Verify select command order matches datagrid column order.
            For Each row In _ds.Tables("ContactsNeeded").Rows
                FrmMain.DataGridView.Rows.Add(row.ItemArray)
            Next
        End Sub

        Public Sub AddContactsToDatasetAccounts()
            Dim bFirstSuccess As Boolean = True
            If (IsNothing(_ds)) Then
                Return
            End If
            If (_ds.Tables.Count = 1) Then
                Dim updateString As StringBuilder = New StringBuilder
                updateString.Append("UPDATE DeferredPaymentAgreements SET DeferredPaymentAgreements.ContactAdded = True WHERE")

                'Add the contact for each account in the dataset.
                For Each row As DataRow In _ds.Tables("ContactsNeeded").Rows
                    'Collect Account Number, DPA Type, Method of delivery and where it was sent.
                    ' Then generate the contact string.
                    Dim rsKey As Integer = row("Key")
                    Dim accountNumber As String = row("AccountNumber")
                    Dim contactInfo As ContactStruct = New ContactStruct()
                    contactInfo.DeliveryMethod = row("DeliveryMethod")
                    contactInfo.DPAType = row("DPAType")
                    contactInfo.SentTo = row("SentTo")
                    Dim contactString As String = GenerateContactString(contactInfo)

                    'Run the script to add the contact
                    If (RunScript(accountNumber, contactString)) Then
                        If (bFirstSuccess) Then
                            updateString.Append(" [DeferredPaymentAgreements].[Key]=" & rsKey.ToString)
                            bFirstSuccess = False 'No longer the first success
                        Else
                            updateString.Append(" OR [DeferredPaymentAgreements].[Key]=" & rsKey.ToString)
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
            contactString = String.Format("{0}ed {1} DPA to {2}", contactInfo.DeliveryMethod, contactInfo.DPAType, contactInfo.SentTo)

            Return contactString
        End Function
    End Module
End Namespace