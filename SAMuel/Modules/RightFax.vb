Imports System
Imports System.IO

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
                       receipiantFax As String, path_to_doc As String) As RFCOMAPILib.Fax
        Dim newFax As RFCOMAPILib.Fax

        newFax = objRightFax.CreateObject(RFCOMAPILib.CreateObjectType.coFax)

        newFax.ToName = receipiantName
        newFax.ToFaxNumber = receipiantFax
        newFax.Attachments.Add(path_to_doc, False) ' false = don't delete file after faxing
        Return newFax
    End Function

    Sub SendFax(ByRef fax As RFCOMAPILib.Fax)
        fax.Send()
    End Sub

    Sub MoveFaxedFile(path_to_doc As String)
        Dim destination As String = My.Settings.savePath + "Faxed/" + Path.GetFileNameWithoutExtension(path_to_doc)
        File.Move(path_to_doc, destination)
    End Sub

End Module
