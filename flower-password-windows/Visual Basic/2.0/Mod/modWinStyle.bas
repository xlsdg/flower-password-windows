Attribute VB_Name = "modWinStyle"
Option Explicit

Private Const GWL_STYLE = (-16)

Private Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME

Private Declare Function GetWindowLong _
                Lib "user32" _
                Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                        ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Public Sub SetWinStyle(ByVal WinhWnd As Long)
    SetWindowLong WinhWnd, GWL_STYLE, GetWindowLong(WinhWnd, GWL_STYLE) And Not WS_CAPTION

End Sub
