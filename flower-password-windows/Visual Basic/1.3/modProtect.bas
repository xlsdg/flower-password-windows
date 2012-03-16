Attribute VB_Name = "modProtect"
Option Explicit

Private Declare Function GetWindowLong _
                Lib "user32" _
                Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                        ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong _
                Lib "user32" _
                Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Private Declare Function CallWindowProc _
                Lib "user32" _
                Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                         ByVal hwnd As Long, _
                                         ByVal Msg As Long, _
                                         ByVal wParam As Long, _
                                         ByVal lParam As Long) As Long

Private Const GWL_WNDPROC = -4

Private Const EM_GETPASSWORDCHAR = &HD2

Private Const WM_GETTEXT = 13

Private Const WM_PASTE = &H302

Private Const WM_COPY = &H301

Private Const WM_RBUTTONDOWN = &H204

Private preWinProc As Long

Private Function wndproc(ByVal hwnd As Long, _
                         ByVal Msg As Long, _
                         ByVal wParam As Long, _
                         ByVal lParam As Long) As Long

    If Msg = WM_GETTEXT Or Msg = EM_GETPASSWORDCHAR Then
        'hWnd = FrmMain.txtKey.hWnd
        wndproc = True
        Exit Function
    ElseIf Msg = WM_COPY Then
        wndproc = True
        Exit Function
    ElseIf Msg = WM_PASTE Then
        wndproc = True
        Exit Function
    ElseIf Msg = WM_RBUTTONDOWN Then
        wndproc = True
        Exit Function

    End If

    wndproc = CallWindowProc(preWinProc, hwnd, Msg, wParam, lParam)

End Function

Public Sub ProtectTextBox(ByVal TextBox_hWnd As Long)
    preWinProc = GetWindowLong(TextBox_hWnd, GWL_WNDPROC)
    Call SetWindowLong(TextBox_hWnd, GWL_WNDPROC, AddressOf wndproc)

End Sub

Public Sub UnProtectTextBox(ByVal TextBox_hWnd As Long)
    Call SetWindowLong(TextBox_hWnd, GWL_WNDPROC, preWinProc)

End Sub
