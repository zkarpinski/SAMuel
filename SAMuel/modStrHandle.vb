Imports System.Text.RegularExpressions

Module modStrHandle
    Function regexAccount(ByVal str As String) As String
        Dim _reg As Regex = New Regex("\d{5}-\d{5}")
        Dim m As Match = _reg.Match(str)
        If m.Success Then
            regexAccount = m.Value
        Else
            regexAccount = "99999-99999"
        End If
    End Function
End Module
