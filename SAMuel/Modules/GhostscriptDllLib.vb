Option Explicit On
Imports System.Runtime.InteropServices

'http://stackoverflow.com/questions/16929383/simple-vb-net-wrapper-for-ghostscript-dll
Module GhostscriptDllLib

    Private Declare Function gsapi_new_instance Lib "gsdll32.dll" _
      (ByRef instance As IntPtr, _
      ByVal caller_handle As IntPtr) As Integer

    Private Declare Function gsapi_set_stdio Lib "gsdll32.dll" _
      (ByVal instance As IntPtr, _
      ByVal gsdll_stdin As StdIOCallBack, _
      ByVal gsdll_stdout As StdIOCallBack, _
      ByVal gsdll_stderr As StdIOCallBack) As Integer

    Private Declare Function gsapi_init_with_args Lib "gsdll32.dll" _
      (ByVal instance As IntPtr, _
      ByVal argc As Integer, _
      <MarshalAs(UnmanagedType.LPArray, ArraySubType:=UnmanagedType.LPStr)> _
      ByVal argv() As String) As Integer

    Private Declare Function gsapi_exit Lib "gsdll32.dll" _
      (ByVal instance As IntPtr) As Integer

    Private Declare Sub gsapi_delete_instance Lib "gsdll32.dll" _
      (ByVal instance As IntPtr)

    '--- Run Ghostscript with specified arguments

    Public Function RunGS(ByVal ParamArray Args() As String) As Boolean

        Dim InstanceHndl As IntPtr
        Dim NumArgs As Integer
        Dim StdErrCallback As StdIOCallBack
        Dim StdInCallback As StdIOCallBack
        Dim StdOutCallback As StdIOCallBack

        NumArgs = Args.Count

        StdInCallback = AddressOf InOutErrCallBack
        StdOutCallback = AddressOf InOutErrCallBack
        StdErrCallback = AddressOf InOutErrCallBack

        '--- Shift arguments to begin at index 1 (Ghostscript requirement)

        ReDim Preserve Args(NumArgs)
        System.Array.Copy(Args, 0, Args, 1, NumArgs)

        '--- Start a new Ghostscript instance

        If gsapi_new_instance(InstanceHndl, 0) <> 0 Then
            Return False
            Exit Function
        End If

        '--- Set up dummy callbacks

        gsapi_set_stdio(InstanceHndl, StdInCallback, StdOutCallback, StdErrCallback)

        '--- Run Ghostscript using specified arguments

        gsapi_init_with_args(InstanceHndl, NumArgs + 1, Args)

        '--- Exit Ghostscript

        gsapi_exit(InstanceHndl)

        '--- Delete instance

        gsapi_delete_instance(InstanceHndl)

        Return True

    End Function

    '--- Delegate function for callbacks

    Private Delegate Function StdIOCallBack(ByVal handle As IntPtr, _
      ByVal Strz As IntPtr, ByVal Bytes As Integer) As Integer

    '--- Dummy callback for standard input, standard output, and errors

    Private Function InOutErrCallBack(ByVal handle As IntPtr, _
      ByVal Strz As IntPtr, ByVal Bytes As Integer) As Integer

        Return 0

    End Function

End Module