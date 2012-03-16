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

Public Type POINTAPI

    X As Long
    Y As Long

End Type

Private Declare Function WindowFromPoint _
                Lib "user32.dll" (ByVal xPoint As Long, _
                                  ByVal yPoint As Long) As Long

Public Function GetDesktopWindowRect(ByRef Rct As RECT) As Long

    Dim MousePos As POINTAPI

    GetCursorPos MousePos

    Dim WinHandle As Long

    WinHandle = WindowFromPoint(MousePos.X, MousePos.Y)

    Dim execute As Integer

    execute = GetWindowRect(WinHandle, Rct)

    If execute = 0 Then
        GetDesktopWindowRect = 0
    Else
        GetDesktopWindowRect = WinHandle

    End If

End Function

Public Function GetMousePoint(ByRef Pos As POINTAPI) As Long
    GetCursorPos Pos
    GetMousePoint = WindowFromPoint(Pos.X, Pos.Y)

End Function
