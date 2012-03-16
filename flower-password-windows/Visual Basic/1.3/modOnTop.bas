Attribute VB_Name = "modOnTop"
Option Explicit

Private Declare Function SetWindowPos _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal hWndInsertAfter As Long, _
                              ByVal x As Long, _
                              ByVal y As Long, _
                              ByVal cx As Long, _
                              ByVal cy As Long, _
                              ByVal wFlags As Long) As Long

Private Const SWP_NOMOVE = &H2

Private Const SWP_NOSIZE = &H1

Private Const HWND_TOPMOST = -1

Private Const HWND_NOTOPMOST = -2

Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Sub SetWinOnTop(ByVal WinHwnd As Long)
    SetWindowPos WinHwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS

End Sub

Public Sub UnSetWinOnTop(ByVal WinHwnd As Long)
    SetWindowPos WinHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS

End Sub
