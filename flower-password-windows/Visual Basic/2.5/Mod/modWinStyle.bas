Attribute VB_Name = "modWinStyle"
Option Explicit

Private Const GWL_STYLE = (-16)

Private Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME

Private Declare Function GetWindowLong _
                Lib "user32.dll" _
                Alias "GetWindowLongA" (ByVal Hwnd As Long, _
                                        ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong _
                Lib "user32.dll" _
                Alias "SetWindowLongA" (ByVal Hwnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Public Function SetWinStyle(ByVal WinHwnd As Long) As Long
    SetWinStyle = SetWindowLong(WinHwnd, GWL_STYLE, GetWindowLong(WinHwnd, GWL_STYLE) And Not WS_CAPTION)

End Function
