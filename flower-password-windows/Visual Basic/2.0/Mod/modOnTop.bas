Attribute VB_Name = "modOnTop"
Option Explicit

Private Declare Function SetWindowPos _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal hWndInsertAfter As Long, _
                              ByVal X As Long, _
                              ByVal Y As Long, _
                              ByVal cx As Long, _
                              ByVal cy As Long, _
                              ByVal wFlags As Long) As Long

Private Const SWP_NOMOVE = &H2

Private Const SWP_NOSIZE = &H1

Private Const HWND_TOPMOST = -1

Private Const HWND_NOTOPMOST = -2

Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Sub SetWinOnTop(ByVal WinhWnd As Long)
    SetWindowPos WinhWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS

End Sub

Public Sub UnSetWinOnTop(ByVal WinhWnd As Long)
    SetWindowPos WinhWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS

End Sub
