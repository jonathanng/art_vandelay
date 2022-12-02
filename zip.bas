Option Explicit

Private Const ZIP_LOCATION = "C:\...\7za"

Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As LongPtr
Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As LongPtr, lpExitCode As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = &H400
Private Const STILL_ACTIVE = &H103


Public Sub unzip(ByVal destination As String, zipFileName As String)

    'x  = extract with full paths
    'o  = output directory
    'ao = overwrite
    shellAndWait path:=ZIP_LOCATION & " x " & zipFileName & " -o" & destination & " -ao", windowsState:=vbHide

End Sub


Public Sub zip(ByVal files As String, ByVal zipFileName As String)

    'a  = add files to archive
    'mx5 = compression level 5
    shellAndWait path:=ZIP_LOCATION & " a " & zipFileName & " " & files & " -mx5", windowsState:=vbHide

End Sub


Private Sub shellAndWait(ByVal path As String, Optional windowsState)

'    Dim wsh As Object
'    Set wsh = VBA.CreateObject("WScript.Shell")
'    Dim waitOnReturn As Boolean: waitOnReturn = True
'    Dim windowStyle As Integer: windowStyle = 1
'
'    wsh.Run path, windowStyle, waitOnReturn

    Dim hProg As Long, hProcess As LongPtr, ExitCode As Long
    If IsMissing(windowsState) Then windowsState = 1                    'fill in the missing parameter
    hProg = Shell(path, windowsState)                                   'execute the program
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, hProg)     'hProg is a "process ID under Win32. used to get the process handle

    Do
        GetExitCodeProcess hProcess, ExitCode                           'populate Exitcode variable
        DoEvents
    Loop While ExitCode = STILL_ACTIVE

End Sub

