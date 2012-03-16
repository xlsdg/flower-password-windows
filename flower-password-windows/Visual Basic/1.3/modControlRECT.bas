Attribute VB_Name = "modControlRECT"
Option Explicit

Private Declare Function GetWindowRect _
                Lib "user32" (ByVal hwnd As Long, _
                              lpRect As RECT) As Long

Public Type RECT

    Left As Long
    Top As Long
    Right As Long
    Bottom As Long

End Type

Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long

Private Type POINTAPI

    x As Long
    y As Long

End Type

Private Declare Function WindowFromPoint _
                Lib "user32.dll" (ByVal xPoint As Long, _
                                  ByVal yPoint As Long) As Long
'Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Public Function GetDesktopWindowRect(ByRef Rct As RECT) As Long

    Dim MousePos As POINTAPI

    GetCursorPos MousePos

    Dim WinHandle As Long

    WinHandle = WindowFromPoint(MousePos.x, MousePos.y)

    'Dim lpClassName As String
    'lpClassName = Space$(256)
    'TEXTBOX Left$(lpClassName, GetClassName(WinHandle, lpClassName, 256))
    Dim execute As Integer

    execute = GetWindowRect(WinHandle, Rct)

    If execute = 0 Then
        GetDesktopWindowRect = ""
    Else
        GetDesktopWindowRect = WinHandle

    End If

End Function
