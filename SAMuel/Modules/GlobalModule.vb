Imports System
Imports System.IO
Imports System.Text.RegularExpressions

Module GlobalModule

    ''' <summary>
    ''' Logs actions and events from within SAMuel
    ''' </summary>
    ''' <param name="actionCode">
    ''' Optional action code to log a predefined event.
    ''' </param>
    ''' <param name="action">
    ''' Optional description of the action.
    ''' </param>
    Public Sub LogAction(Optional actionCode As Integer = 0, Optional action As String = vbNullString)
        Dim logFile As String = My.Settings.savePath + "SAMuel.log"
        Dim sw As StreamWriter
        Dim logAction As String

        ' Create the log file if it doesn't exist 
        If File.Exists(logFile) = False Then
            sw = File.CreateText(logFile)
            sw.Flush()
            sw.Close()
        End If

        'Choose the appropriate log output
        sw = File.AppendText(logFile)
        Select Case actionCode
            Case 0 'Custom log or error
                If action IsNot vbNullString Then
                    logAction = action
                Else
                    logAction = "Unknown action occured."
                End If
            Case 1 'Faxed File
                logAction = action + " was faxed."
            Case 2 'Save path changed
                logAction = "Default save path changed to " + action
            Case Else
                logAction = "Unknown action code used."
        End Select

        sw.WriteLine(Now + ": " + logAction)

        'Close the log file
        sw.Flush()
        sw.Close()
    End Sub

    Function RegexAccount(ByVal str As String) As String
        Dim _reg As Regex = New Regex("\d{5}-\d{5}")
        Dim m As Match = _reg.Match(str)
        If m.Success Then
            Return m.Value
        Else
            Return "ACC# NOT FOUND"
        End If
    End Function

    Function RegexCustomer(ByVal str As String) As String
        Dim _reg As Regex = New Regex("\d{9}")
        Dim m As Match = _reg.Match(str)
        If m.Success Then
            Return m.Value
        Else
            Return "CUST# NOT FOUND"
        End If
    End Function

End Module
