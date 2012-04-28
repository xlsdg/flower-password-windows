Attribute VB_Name = "modProtect"
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

Private Declare Function GetWindowLong _
                Lib "user32.dll" _
                Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                        ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong _
                Lib "user32.dll" _
                Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Private Declare Function CallWindowProc _
                Lib "user32.dll" _
                Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                         ByVal hwnd As Long, _
                                         ByVal Msg As Long, _
                                         ByVal wParam As Long, _
                                         ByVal lParam As Long) As Long

Private Const GWL_WNDPROC = (-4)

Private Const EM_GETPASSWORDCHAR = &HD2

Private Const WM_GETTEXT = &HD

Private Const WM_PASTE = &H302

Private Const WM_RBUTTONDOWN = &H204

Private Const WM_SETTEXT = &HC

Private Const WM_GETTEXTLENGTH = &HE

Private Const WM_COPY = &H301

Private Const WM_COPYDATA = &H4A

Private Const WM_CUT = &H300

Private Const EM_REPLACESEL = &HC2

Private preEditProc As Long

Public Function ProtectTextBox(ByVal TextBox_hWnd As Long) As Long
    preEditProc = GetWindowLong(TextBox_hWnd, GWL_WNDPROC)
    ProtectTextBox = SetWindowLong(TextBox_hWnd, GWL_WNDPROC, AddressOf EditProc)

End Function

Public Function UnProtectTextBox(ByVal TextBox_hWnd As Long) As Long
    UnProtectTextBox = SetWindowLong(TextBox_hWnd, GWL_WNDPROC, preEditProc)

End Function

Private Function EditProc(ByVal hwnd As Long, _
                          ByVal Msg As Long, _
                          ByVal wParam As Long, _
                          ByVal lParam As Long) As Long

    If Msg = WM_GETTEXT Or Msg = EM_GETPASSWORDCHAR Or Msg = WM_COPY Or Msg = WM_PASTE Or Msg = WM_RBUTTONDOWN Or Msg = WM_CUT Then
        'hwnd = FrmMain.txtKey.hwnd
        EditProc = True
        Exit Function

    End If

    EditProc = CallWindowProc(preEditProc, hwnd, Msg, wParam, lParam)

End Function
