Attribute VB_Name = "modMouse"
Option Explicit

Private Declare Function SetWindowsHookEx _
                Lib "user32" _
                Alias "SetWindowsHookExA" (ByVal idHook As Long, _
                                           ByVal lpfn As Long, _
                                           ByVal hmod As Long, _
                                           ByVal dwThreadId As Long) As Long

Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Private Declare Function CallNextHookEx _
                Lib "user32" (ByVal hHook As Long, _
                              ByVal ncode As Long, _
                              ByVal wParam As Long, _
                              lParam As Any) As Long

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (lpvDest As Any, _
                                       ByVal lpvSource As Long, _
                                       ByVal cbCopy As Long)

Private Type MOUSEMSGS

    X As Long
    Y As Long
    a As Long
    b As Long
    time As Long

End Type

Private Const WH_MOUSE_LL = 14

Private Const HC_ACTION = 0

Private Const WM_MOUSEMOVE = &H200

Private Const WM_LBUTTONDOWN = &H201

Private Const WM_LBUTTONUP = &H202

Private Const WM_LBUTTONDBLCLK = &H203

Private Const WM_RBUTTONDOWN = &H204

Private Const WM_RBUTTONUP = &H205

Private Const WM_RBUTTONDBLCLK = &H206

Private Const WM_MBUTTONDOWN = &H207

Private Const WM_MBUTTONUP = &H208

Private Const WM_MBUTTONDBLCLK = &H209

Private Const WM_MOUSEACTIVATE = &H21

Private Const WM_MOUSEFIRST = &H200

Private Const WM_MOUSELAST = &H209

Private Const WM_MOUSEWHEEL = &H20A

Private MouseMsg As MOUSEMSGS

Private lHook    As Long

Private Function CallMouseHookProc(ByVal code As Long, _
                                   ByVal wParam As Long, _
                                   ByVal lParam As Long) As Long

    If code = HC_ACTION Then
        CopyMemory MouseMsg, lParam, LenB(MouseMsg)

        If (MouseMsg.X * Screen.TwipsPerPixelX) < FrmMain.Left Or (MouseMsg.Y * Screen.TwipsPerPixelY) < FrmMain.Top Or (MouseMsg.X * Screen.TwipsPerPixelX) > (FrmMain.Left + FrmMain.Width) Or (MouseMsg.Y * Screen.TwipsPerPixelY) > (FrmMain.Top + FrmMain.Height) Then
            If wParam = WM_LBUTTONDOWN And FrmMain.Visible Then
                FrmMain.Visible = False
                PostCode Password_Hwnd, getFlowerPassword(FrmMain.txtPassword.Text, FrmMain.txtKey.Text, 16)
            ElseIf (wParam = WM_RBUTTONDOWN Or wParam = WM_MBUTTONDOWN) And FrmMain.Visible Then
                FrmMain.Visible = False

            End If

            'Select Case wParam
            'Case WM_LBUTTONDOWN
            'Case WM_MBUTTONDOWN
            'Case WM_RBUTTONDOWN
            'End Select
        End If

        CallMouseHookProc = 0

    End If

    If code <> 0 Then
        CallMouseHookProc = CallNextHookEx(0, code, wParam, lParam)

    End If

End Function

Public Sub SetMouseHook()
    lHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf CallMouseHookProc, App.hInstance, 0)

End Sub

Public Sub UnSetMouseHook()
    UnhookWindowsHookEx lHook

End Sub
