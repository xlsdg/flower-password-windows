Attribute VB_Name = "mod64bit"
Option Explicit

Private Declare Function GetProcAddress _
                Lib "kernel32.dll" (ByVal hModule As Long, _
                                ByVal lpProcName As String) As Long

Private Declare Function GetModuleHandle _
                Lib "kernel32.dll" _
                Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Declare Function IsWow64Process _
                Lib "kernel32.dll" (ByVal hProc As Long, _
                                bWow64Process As Boolean) As Long

Public Function Is64bit() As Boolean

    Dim handle As Long, bolFunc As Boolean

    ' Assume initially that this is not a Wow64 process
    bolFunc = False
    ' Now check to see if IsWow64Process function exists
    handle = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")

    If handle > 0 Then ' IsWow64Process function exists
        ' Now use the function to determine if
        ' we are running under Wow64
        IsWow64Process GetCurrentProcess(), bolFunc

    End If

    Is64bit = bolFunc

End Function
