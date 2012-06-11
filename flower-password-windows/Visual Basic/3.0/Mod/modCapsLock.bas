Attribute VB_Name = "modCapsLock"
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

Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer

Private Declare Function MapVirtualKey _
                Lib "user32.dll" _
                Alias "MapVirtualKeyA" (ByVal wCode As Long, _
                                        ByVal wMapType As Long) As Long

Private Declare Sub keybd_event _
                Lib "user32.dll" (ByVal bVk As Byte, _
                                  ByVal bScan As Byte, _
                                  ByVal dwFlags As Long, _
                                  ByVal dwExtraInfo As Long)

Private Const KEYEVENTF_EXTENDEDKEY = &H1

Private Const KEYEVENTF_KEYUP = &H2

Public Sub SetCapsLock(ByVal bLock As Boolean)

    Dim Check As Boolean

    Check = CBool(GetKeyState(vbKeyCapital))

    If Check <> bLock Then

        Dim Scancode As Long

        Scancode = MapVirtualKey(vbKeyCapital, 0)
        keybd_event vbKeyCapital, Scancode, 0, 0
        keybd_event vbKeyCapital, Scancode, KEYEVENTF_KEYUP, 0

    End If

End Sub
