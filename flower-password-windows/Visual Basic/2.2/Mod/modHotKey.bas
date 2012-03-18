Attribute VB_Name = "modHotKey"
Option Explicit

Private Declare Function SetWindowLong _
                Lib "user32.dll" _
                Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowLong _
                Lib "user32.dll" _
                Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                        ByVal nIndex As Long) As Long

Private Declare Function CallWindowProc _
                Lib "user32.dll" _
                Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                         ByVal hwnd As Long, _
                                         ByVal Msg As Long, _
                                         ByVal wParam As Long, _
                                         ByVal lParam As Long) As Long

Private Declare Function RegisterHotKey _
                Lib "user32.dll" (ByVal hwnd As Long, _
                                  ByVal id As Long, _
                                  ByVal fsModifiers As Long, _
                                  ByVal vk As Long) As Long

Private Declare Function UnregisterHotKey _
                Lib "user32.dll" (ByVal hwnd As Long, _
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

Private Function HotKeyProc(ByVal hwnd As Long, _
                            ByVal Msg As Long, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long) As Long

    If Msg = WM_HOTKEY Then
        If wParam = idHotKey Then

            Dim lp As taLong, i2 As t2Int

            lp.ll = lParam
            LSet i2 = lp

            If (i2.lWord = Modifiers) And i2.hword = uVirtKey Then
                If Not FrmMain.Visible Then

                    Dim caretpos As POINTAPI, mousepos As POINTAPI, rects As Rect, caretpos_hWnd As Long, rect_hWnd As Long

                    caretpos_hWnd = GetDesktopWindowCaretPos(caretpos)
                    rect_hWnd = GetDesktopWindowRect(rects, mousepos)

                    If caretpos_hWnd = rect_hWnd And rect_hWnd <> 0 Then
                        If (rects.Bottom - caretpos.y) < 30 Then
                            Password_Hwnd = rect_hWnd
                            'SetWinByPoint FrmMain.hwnd, Rects.Left * Screen.TwipsPerPixelX, Rects.Bottom * Screen.TwipsPerPixelY
                            FrmMain.Top = rects.Bottom * Screen.TwipsPerPixelY
                            FrmMain.Left = rects.Left * Screen.TwipsPerPixelX
                        ElseIf (mousepos.y - caretpos.y) < 50 Then
                            Password_Hwnd = caretpos_hWnd
                            'SetWinByPoint FrmMain.hwnd, MousePos.x * Screen.TwipsPerPixelX, (MousePos.y + 17) * Screen.TwipsPerPixelY
                            FrmMain.Top = (caretpos.y + 17) * Screen.TwipsPerPixelY
                            FrmMain.Left = caretpos.x * Screen.TwipsPerPixelX
                        Else
                            Password_Hwnd = rect_hWnd
                            'SetWinByPoint FrmMain.hwnd, MousePos.x * Screen.TwipsPerPixelX, (MousePos.y + 10) * Screen.TwipsPerPixelY
                            FrmMain.Top = (mousepos.y + 10) * Screen.TwipsPerPixelY
                            FrmMain.Left = mousepos.x * Screen.TwipsPerPixelX

                        End If

                    ElseIf caretpos_hWnd <> 0 Then
                        Password_Hwnd = caretpos_hWnd
                        'SetWinByPoint FrmMain.hwnd, MousePos.x * Screen.TwipsPerPixelX, (MousePos.y + 17) * Screen.TwipsPerPixelY
                        FrmMain.Top = (caretpos.y + 17) * Screen.TwipsPerPixelY
                        FrmMain.Left = caretpos.x * Screen.TwipsPerPixelX
                    Else
                        Password_Hwnd = rect_hWnd
                        'SetWinByPoint FrmMain.hwnd, MousePos.x * Screen.TwipsPerPixelX, (MousePos.y + 10) * Screen.TwipsPerPixelY
                        FrmMain.Top = (mousepos.y + 10) * Screen.TwipsPerPixelY
                        FrmMain.Left = mousepos.x * Screen.TwipsPerPixelX

                    End If

                    FrmMain.Visible = True
                    FrmMain.txtPassword.SetFocus

                End If

            End If

        End If

    End If

    HotKeyProc = CallWindowProc(preWinProc, hwnd, Msg, wParam, lParam)

End Function

Public Function SetHotKey(ByVal WinhWnd As Long) As Long
    preWinProc = GetWindowLong(WinhWnd, GWL_WNDPROC)
    SetWindowLong WinhWnd, GWL_WNDPROC, AddressOf HotKeyProc
    idHotKey = 1
    Modifiers = MOD_WIN
    uVirtKey = vbKeyS
    SetHotKey = RegisterHotKey(WinhWnd, idHotKey, Modifiers, uVirtKey)

End Function

Public Function UnSetHotKey(ByVal WinhWnd As Long) As Long
    SetWindowLong WinhWnd, GWL_WNDPROC, preWinProc
    UnSetHotKey = UnregisterHotKey(WinhWnd, uVirtKey)

End Function
