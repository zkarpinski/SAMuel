Imports System
Imports System.IO
Imports System.Text.RegularExpressions

Module GlobalModule
    ' Global Variables (Loaded from .ini)

    ''' <summary>
    ''' Logs actions and events from within SAMuel
    ''' </summary>
    ''' <param name="actionCode">
    ''' Optional action code to log a predefined event.
    ''' 0: Custom Log  1: Faxed file  2: Save Path Changed  3:Email Processed  50: Email Skipped  51: Unknown Filetype  98: SAM email error  99: Undesired state reached
    ''' 
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
            Case 1  'Faxed File
                logAction = "Notice  : " + action + " was faxed."
            Case 2  'Save path changed
                logAction = "Notice  : Default save path changed to " + action
            Case 3  'Email Processed
                logAction = "Notice  : Email - " + action + " was processed"
            Case 50 'Email Skipped
                logAction = "Warning : Email - " + action
            Case 51 'Unknown Filetype'
                logAction = "Unknown filetype (" + action + ")."
            Case 98 'SAM email error
                logAction = "Error 98: Email Error: " + action
            Case 99 'Reached undesired state
                logAction = "Error 99: Undesired state reached. " + action
            Case Else
                logAction = "ERROR XX: Unknown action code used."
        End Select

        sw.WriteLine(Now + "| " + logAction)

        'Close the log file
        sw.Flush()
        sw.Close()
    End Sub

    ''' <summary>
    ''' Calls regex.
    ''' </summary>
    ''' <param name="str">String to process.</param>
    ''' <param name="format">Regex format as a string.</param>
    ''' <returns>String of the found match or X if no match.</returns>
    ''' <remarks></remarks>
    Function RegexAcc(ByVal str As String, ByVal format As String) As String
        Dim _reg As Regex = New Regex(format)
        Dim m As Match = _reg.Match(str)
        If m.Success Then
            Return m.Value
        Else
            Return "X"
        End If
    End Function

    ''' <summary>
    ''' Creates the output folders used by SAMuel if they do not exist
    ''' </summary>
    Sub InitOutputFolders()
        Dim parentPath As String = My.Settings.savePath
        Dim OutputFolders As Array = {parentPath, parentPath + "tiffs\", parentPath + "emails\", parentPath + "faxed\", parentPath + "converted\"}

        'Create the folder if it does not exist
        For Each value In OutputFolders
            GlobalModule.CheckFolder(value)
        Next
    End Sub

    ''' <summary>
    ''' Creates a folder if the path does not exist
    ''' </summary>
    Public Sub CheckFolder(ByVal path As String)
        If (Not System.IO.Directory.Exists(path)) Then
            System.IO.Directory.CreateDirectory(path)
        End If
    End Sub

    ''' <summary>
    ''' Generates a random string
    ''' </summary>
    ''' <param name="length"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function RandomString(ByVal length As Integer) As String
        Dim random As New Random()
        Dim charOutput As Char() = New Char(length - 1) {}
        For i As Integer = 0 To length - 1
            Dim selector As Integer = random.[Next](65, 101)
            If selector > 90 Then
                selector -= 43
            End If
            charOutput(i) = Convert.ToChar(selector)
        Next
        Return New String(charOutput)
    End Function

    ''' <summary>
    '''  Strips invalid characters from a string.
    ''' </summary>
    ''' <param name="strIn"></param>
    ''' <returns>Clean string</returns>
    ''' <remarks>http://msdn.microsoft.com/en-us/library/844skk0h</remarks>
    Function CleanInput(strIn As String) As String
        ' Replace invalid characters with empty strings.
        ' Allows all punctuation, letters, numbers, @ symbol, hyphen and space.
        Return Regex.Replace(strIn, "[^\w\s\.@-]", "")
    End Function

End Module
