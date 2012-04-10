Attribute VB_Name = "modWinStyle"
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

Private Const GWL_STYLE = (-16)

Private Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME

Private Declare Function GetWindowLong _
                Lib "user32.dll" _
                Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                        ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong _
                Lib "user32.dll" _
                Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                        ByVal nIndex As Long, _
                                        ByVal dwNewLong As Long) As Long

Public Function SetWinStyle(ByVal WinHwnd As Long) As Long
    SetWinStyle = SetWindowLong(WinHwnd, GWL_STYLE, GetWindowLong(WinHwnd, GWL_STYLE) And Not WS_CAPTION)

End Function
