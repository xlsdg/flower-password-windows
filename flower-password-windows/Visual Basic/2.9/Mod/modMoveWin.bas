Attribute VB_Name = "modMoveWin"
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

Private Declare Function ReleaseCapture Lib "user32.dll" () As Long

Private Declare Function SendMessage _
                Lib "user32.dll" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Private Const WM_SYSCOMMAND = &H112

Private Const SC_MOVE = &HF010&

Private Const WM_NCLBUTTONDOWN = &HA1

Private Const HTCAPTION = 2

Public Function SetWinMove(ByVal WinHwnd As Long) As Long
    ReleaseCapture
    SetWinMove = SendMessage(WinHwnd, WM_SYSCOMMAND, SC_MOVE + HTCAPTION, 0&)

    'SendMessage WinHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Function
