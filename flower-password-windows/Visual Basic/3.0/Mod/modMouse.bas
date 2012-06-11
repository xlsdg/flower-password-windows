Attribute VB_Name = "modMouse"
'*****************************************************************
' Copyright (c) 2011-2012 FlowerPassword.com All rights reserved.
'      Author : xLsDg @ Xiao Lu Software Development Group
'        Blog : http://hi.baidu.com/xlsdg
'          QQ : 4 4 7 4 0 5 7 4 0
'     Version : 1 . 0 . 0 . 0
'        Date : 2 0 1 2 / 0 4 / 0 7
' Description :
'     History :
'*****************************************************************
Option Explicit

Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long

Private Declare Function WindowFromPoint _
                Lib "user32.dll" (ByVal xPoint As Long, _
                                  ByVal yPoint As Long) As Long

Private Type RECT

    Left As Long
    Top As Long
    Right As Long
    Bottom As Long

End Type

Private Declare Function GetWindowRect _
                Lib "user32.dll" (ByVal hwnd As Long, _
                                  lpRect As RECT) As Long

Private Declare Function SetWindowsHookEx _
                Lib "user32.dll" _
                Alias "SetWindowsHookExA" (ByVal idHook As Long, _
                                           ByVal lpfn As Long, _
                                           ByVal hmod As Long, _
                                           ByVal dwThreadId As Long) As Long

Private Declare Function UnhookWindowsHookEx _
                Lib "user32.dll" (ByVal hHook As Long) As Long

Private Declare Function CallNextHookEx _
                Lib "user32.dll" (ByVal hHook As Long, _
                                  ByVal ncode As Long, _
                                  ByVal wParam As Long, _
                                  lParam As Any) As Long

Private Declare Sub CopyMemory _
                Lib "kernel32.dll" _
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

Public Function SetMouseHook() As Long
    lHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf CallMouseHookProc, App.hInstance, 0)
    SetMouseHook = lHook

End Function

Public Function UnSetMouseHook() As Long
    UnSetMouseHook = UnhookWindowsHookEx(lHook)

End Function

Private Function CallMouseHookProc(ByVal code As Long, _
                                   ByVal wParam As Long, _
                                   ByVal lParam As Long) As Long

    If code = HC_ACTION Then
        If FrmMain.Visible Then

            Dim frmRec As RECT

            Select Case wParam

                Case WM_LBUTTONDOWN '左键按下
                    CopyMemory MouseMsg, lParam, LenB(MouseMsg)
                    GetWindowRect FrmMain.hwnd, frmRec

                    If frmRec.Left <= MouseMsg.X And MouseMsg.X <= frmRec.Right And frmRec.Top <= MouseMsg.Y And MouseMsg.Y <= (frmRec.Bottom + 130) Then
                    Else
                        FrmMain.FrmHide True, False

                    End If

                    CallMouseHookProc = 0

                Case WM_MBUTTONDOWN '中键按下

                    'CopyMemory MouseMsg, lParam, LenB(MouseMsg)
                    'GetWindowRect FrmMain.hwnd, frmRec
                    'If frmRec.Left <= MouseMsg.X And MouseMsg.X <= frmRec.Right And frmRec.Top <= MouseMsg.Y And MouseMsg.Y <= frmRec.Bottom Then
                    'Else
                    'FrmMain.FrmHide False, False
                    'End If
                    'CallMouseHookProc = 0
                Case WM_RBUTTONDOWN '右键按下

                    'CopyMemory MouseMsg, lParam, LenB(MouseMsg)
                    'GetWindowRect FrmMain.hwnd, frmRec
                    'If frmRec.Left <= MouseMsg.X And MouseMsg.X <= frmRec.Right And frmRec.Top <= MouseMsg.Y And MouseMsg.Y <= frmRec.Bottom Then
                    'Else
                    'FrmMain.FrmHide False, False
                    'End If
                    'CallMouseHookProc = 0
            End Select

        End If

    End If

    If code <> 0 Then
        CallMouseHookProc = CallNextHookEx(0, code, wParam, lParam)

    End If

End Function

