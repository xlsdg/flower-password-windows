Attribute VB_Name = "modOnTop"
Option Explicit

Private Declare Function SetWindowPos _
                Lib "user32.dll" (ByVal hwnd As Long, _
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

Public Function SetWinOnTop(ByVal WinhWnd As Long) As Long
    SetWinOnTop = SetWindowPos(WinhWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

End Function

Public Function UnSetWinOnTop(ByVal WinhWnd As Long) As Long
    UnSetWinOnTop = SetWindowPos(WinhWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)

End Function
