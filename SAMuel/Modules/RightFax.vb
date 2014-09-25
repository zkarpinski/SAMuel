
Imports System.IO

Namespace Modules

    Module RightFax
        Function ConnectToServer(rfServer As String, rfUser As String, _
                                 useNT As Boolean) As RFCOMAPILib.FaxServer
            Dim objRightFax As RFCOMAPILib.FaxServer
            objRightFax = New RFCOMAPILib.FaxServer

            objRightFax.ServerName = rfServer
            objRightFax.Protocol = 4
            objRightFax.AuthorizationUserID = rfUser
            objRightFax.UseNTAuthentication = useNT

            Return objRightFax
        End Function

        Function CreateFax(ByRef objRightFax As RFCOMAPILib.FaxServer, receipiantName As String, _
                           receipiantFax As String, pathToDoc As String, Optional ByVal coverSheetNotes As String = vbNullString) As RFCOMAPILib.Fax
            Dim newFax As RFCOMAPILib.Fax

            newFax = objRightFax.CreateObject(RFCOMAPILib.CreateObjectType.coFax)

            newFax.ToName = receipiantName
            newFax.ToFaxNumber = receipiantFax
            'newFax.CoverSheetNotes(0) = coverSheetNotes
            newFax.UserComments = "Sent via SAMuel."
            newFax.Attachments.Add(pathToDoc, False) ' false = don't delete file after faxing

            Return newFax
        End Function

        Sub SendFax(ByRef fax As RFCOMAPILib.Fax)
            fax.Send()
        End Sub

        Sub MoveFaxedFile(pathToDoc As String)
            Dim destination As String = GlobalModule.FAXED_FOLDER + Path.GetFileName(pathToDoc)
            File.Move(pathToDoc, destination)
        End Sub

    End Module
End Namespace