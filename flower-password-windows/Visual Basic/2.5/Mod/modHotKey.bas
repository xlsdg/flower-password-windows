Attribute VB_Name = "modHotKey"
Option Explicit

Private Declare Function MoveWindow _
                Lib "user32.dll" (ByVal hWnd As Long, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal nWidth As Long, _
                                  ByVal nHeight As Long, _
                                  ByVal bRepaint As Long) As Long

Private Declare Function SetWindowLong _
                Lib "user32.dll" _
                Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowLong _
                Lib "user32.dll" _
                Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                        ByVal nIndex As Long) As Long

Private Declare Function CallWindowProc _
                Lib "user32.dll" _
                Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                         ByVal hWnd As Long, _
                                         ByVal msg As Long, _
                                         ByVal wParam As Long, _
                                         ByVal lParam As Long) As Long

Private Declare Function RegisterHotKey _
                Lib "user32.dll" (ByVal hWnd As Long, _
                                  ByVal id As Long, _
                                  ByVal fsModifiers As Long, _
                                  ByVal vk As Long) As Long

Private Declare Function UnregisterHotKey _
                Lib "user32.dll" (ByVal hWnd As Long, _
                                  ByVal id As Long) As Long

Private Const WM_HOTKEY = &H312

Private Const MOD_ALT = &H1

Private Const MOD_CONTROL = &H2

Private Const MOD_SHIFT = &H4

Private Const MOD_WIN = &H8

Private Const GWL_WNDPROC = (-4)

Private preWinProc As Long

Private Modifiers  As Long, uVirtKey As Long, idHotKey As Long

Private Type taLong

    ll As Long

End Type

Private Type t2Int

    lWord As Integer
    hword As Integer

End Type

Public Function SetHotKey(ByVal WinHwnd As Long) As Long
    preWinProc = GetWindowLong(WinHwnd, GWL_WNDPROC)
    SetWindowLong WinHwnd, GWL_WNDPROC, AddressOf HotKeyProc
    idHotKey = 1
    Modifiers = MOD_WIN
    uVirtKey = vbKeyS
    SetHotKey = RegisterHotKey(WinHwnd, idHotKey, Modifiers, uVirtKey)

End Function

Public Function UnSetHotKey(ByVal WinHwnd As Long) As Long
    SetWindowLong WinHwnd, GWL_WNDPROC, preWinProc
    UnSetHotKey = UnregisterHotKey(WinHwnd, uVirtKey)

End Function

Private Function HotKeyProc(ByVal hWnd As Long, _
                            ByVal msg As Long, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long) As Long

    If msg = WM_HOTKEY Then
        If wParam = idHotKey Then

            Dim lp As taLong, i2 As t2Int

            lp.ll = lParam
            LSet i2 = lp

            If (i2.lWord = Modifiers) And i2.hword = uVirtKey Then
                If Not FrmMain.Visible Then

                    Dim X As Long, Y As Long

                    getLocation X, Y
                    SetUrlAsKey Password_Hwnd
                    MoveWindow FrmMain.hWnd, X, Y, FrmMain.Width / Screen.TwipsPerPixelX, FrmMain.Height / Screen.TwipsPerPixelY, False
                    FrmMain.FrmShow

                End If

            End If

        End If

    End If

    HotKeyProc = CallWindowProc(preWinProc, hWnd, msg, wParam, lParam)

End Function


