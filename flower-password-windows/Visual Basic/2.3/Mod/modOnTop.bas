Attribute VB_Name = "modOnTop"
Option Explicit

'这个函数能为窗口指定一个新位置和状态。它也可改变窗口在内部窗口列表中的位置。该函数与DeferWindowPos函数相似，只是它的作用是立即表现出来的（在vb里使用：针对vb窗体，如它们在win32下屏蔽或最小化，则需重设最顶部状态。如有必要，请用一个子类处理模块来重设最顶部状态
Private Declare Function SetWindowPos _
                Lib "user32" (ByVal hwnd As Long, _
                              ByVal hWndInsertAfter As Long, _
                              ByVal x As Long, _
                              ByVal y As Long, _
                              ByVal cx As Long, _
                              ByVal cy As Long, _
                              ByVal wFlags As Long) As Long

Private Const SWP_NOACTIVATE = &H10

Private Const SWP_SHOWWINDOW = &H40

Private Const SWP_NOMOVE = &H2

Private Const SWP_NOSIZE = &H1

Private Const HWND_TOPMOST = -1

Private Const HWND_NOTOPMOST = -2

Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Function SetWinOnTop(ByVal WinHwnd As Long) As Long
    SetWinOnTop = SetWindowPos(WinHwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

End Function

Public Function UnSetWinOnTop(ByVal WinHwnd As Long) As Long
    UnSetWinOnTop = SetWindowPos(WinHwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)

End Function

Public Function SetWinByPoint(ByVal WinHwnd As Long, _
                              ByVal point_x As Long, _
                              ByVal point_y As Long) As Long
    SetWinByPoint = SetWindowPos(WinHwnd, HWND_TOPMOST, point_x, point_y, 0, 0, SWP_NOSIZE Or SWP_SHOWWINDOW)

End Function
